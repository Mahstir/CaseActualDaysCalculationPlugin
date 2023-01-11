using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.PluginTelemetry;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseActualDaysCalculator
{
    public class CaseActualDaysCalculation : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            

            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.  
                Entity entity = (Entity)context.InputParameters["Target"];
                Guid Id = entity.Id;

                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                var status = service.Retrieve("incident", Id, new ColumnSet("statecode"));
                var caseStatus = status.GetAttributeValue<OptionSetValue>("statecode").Value;

                if (caseStatus == 1) { 
                    
                    try
                    {


                        var calendarRules = new List<Entity>();
                        Entity businessClosureCalendar = null;

                        QueryExpression q = new QueryExpression("calendar") { NoLock = true };
                        q.Criteria.AddCondition("type", ConditionOperator.Equal, 2);
                        q.Criteria.AddCondition("name", ConditionOperator.Equal, "Holiday Schedule");


                        EntityCollection businessClosureCalendars = service.RetrieveMultiple(q);

                        if (businessClosureCalendars.Entities.Count > 0)
                        {
                            businessClosureCalendar = businessClosureCalendars.Entities[0];
                            calendarRules = businessClosureCalendar.GetAttributeValue<EntityCollection>("calendarrules").Entities.ToList();

                        }

                        var neededDates = service.Retrieve("incident", Id, new ColumnSet("createdon", "modifiedon"));

                        List<DateTime> allDatesInbetween = new List<DateTime>();
                        var startDate = Convert.ToDateTime(neededDates.Attributes["createdon"]).Date;
                        var endDate = Convert.ToDateTime(neededDates.Attributes["modifiedon"]).Date;


                        for (var date = startDate; date <= endDate; date = date.AddDays(1))
                        {

                            if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                            {
                                allDatesInbetween.Add(date);
                            }

                        }
                       

                        foreach (var holidayDate in calendarRules)
                        {

                            var holiday = Convert.ToDateTime(holidayDate.GetAttributeValue<DateTime>("starttime")).Date;
                            if (allDatesInbetween.Contains(holiday))
                                allDatesInbetween.Remove(holiday);

                        }

                        var actualNumberOfDays = allDatesInbetween.Count();

                        Entity updateEntity = new Entity("incident");
                        updateEntity.Id = entity.Id;
                        updateEntity["new_noofdaystakentoresolve"] = actualNumberOfDays;
                        service.Update(updateEntity);

                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace("FollowUpPlugin: {0}", ex.ToString());
                        throw;
                    }


               }
            }

        }
    }
}
