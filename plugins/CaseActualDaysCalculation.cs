using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;

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

                tracingService.Trace("Getting Case state");
                var caseStatus = entity.GetAttributeValue<OptionSetValue>("statecode").Value;
                tracingService.Trace("Case status is {0}", caseStatus);

                if (caseStatus == 1)
                {

                    try
                    {

                        tracingService.Trace("Retrieving Holidays");
                        var calendarRules = new List<Entity>();
                        Entity businessClosureCalendar = null;

                        QueryExpression q = new QueryExpression("calendar") { NoLock = true };
                        q.Criteria.AddCondition("type", ConditionOperator.Equal, 2);
                        q.Criteria.AddCondition("name", ConditionOperator.Equal, "Bank Holidays");


                        EntityCollection businessClosureCalendars = service.RetrieveMultiple(q);

                        if (businessClosureCalendars.Entities.Count > 0)
                        {
                            tracingService.Trace("Holidays");

                            businessClosureCalendar = businessClosureCalendars.Entities[0];
                            calendarRules = businessClosureCalendar.GetAttributeValue<EntityCollection>("calendarrules").Entities.ToList();

                        }


                        tracingService.Trace("Retrieving received and actual resolved dates");

                        var neededDates = service.Retrieve("incident", Id, new ColumnSet("new_ltc_received_date", "new_ltc_actual_resolved_date", "ticketnumber"));

                        tracingService.Trace("Checking Dates exist for {0}", neededDates.Attributes["ticketnumber"]);
                        if (!neededDates.Contains("new_ltc_actual_resolved_date"))
                        {
                            throw new InvalidPluginExecutionException("Actual Resolved date is required");
                        }

                        tracingService.Trace("Initialising Dates");

                        List<DateTime> allDatesInbetween = new List<DateTime>();

                        var startDate = Convert.ToDateTime(neededDates.Attributes["new_ltc_received_date"]).Date;
                        tracingService.Trace("Start date: {0}", startDate);

                        var endDate = Convert.ToDateTime(neededDates.Attributes["new_ltc_actual_resolved_date"]).Date;
                        tracingService.Trace("Start date: {0}", endDate);


                        tracingService.Trace("Excluding weekends from dates in between");
                        for (var date = startDate; date <= endDate; date = date.AddDays(1))
                        {

                            if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                            {
                                allDatesInbetween.Add(date);
                            }

                        }

                        tracingService.Trace("Exclusing holidays from days in between");
                        foreach (var holidayDate in calendarRules)
                        {

                            var holiday = Convert.ToDateTime(holidayDate.GetAttributeValue<DateTime>("starttime")).Date;
                            if (allDatesInbetween.Contains(holiday))
                                allDatesInbetween.Remove(holiday);

                        }

                        var actualNumberOfDays = allDatesInbetween.Count();

                        tracingService.Trace("Updating Incident");
                        Entity updateEntity = new Entity("incident")
                        {
                            Id = entity.Id
                        };
                        updateEntity["tisski_noofworkingdaystakentoresolve"] = actualNumberOfDays;
                        service.Update(updateEntity);

                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        tracingService.Trace("Number of Days taken to resolve: {0}", ex.ToString());
                        throw new InvalidPluginExecutionException("An error occurred in the CaseActualDaysCalculation plug-in.", ex);
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace("CaseActualDaysCalculationPlugin: {0}", ex.ToString());
                        throw;
                    }


                }
            }

        }
    }
}
