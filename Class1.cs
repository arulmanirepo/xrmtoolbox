using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

public class CalculateCaseResolutionBusinessMinutes : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
        var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

        tracingService.Trace("Plugin execution started");

        tracingService.Trace("Case Id: {0}", context.PrimaryEntityId.ToString());

        if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity incidentResolution)
        {
            if (incidentResolution.Contains("incidentid") && incidentResolution.Contains("actualend"))
            {
                EntityReference incidentRef = incidentResolution.GetAttributeValue<EntityReference>("incidentid");
                DateTime actualEnd = incidentResolution.GetAttributeValue<DateTime>("actualend");

                // Retrieve the incident
                Entity incident = service.Retrieve("incident", incidentRef.Id, new ColumnSet("createdon"));
                if (incident != null && incident.Contains("createdon"))
                {
                    DateTime createdOn = incident.GetAttributeValue<DateTime>("createdon");

                    // Calculate total business minutes
                    //int businessMinutes = GetResolutionMinutes(createdOn, actualEnd, tracingService);
                    int resolutionMinutes = GetResolutionMinutes(createdOn, actualEnd, tracingService);
                    int netResolutionMinutes = GetNetResolutionMinutes(createdOn, actualEnd, tracingService);

                    // Update the incident
                    Entity updateIncident = new Entity("incident")
                    {
                        Id = incident.Id
                    };

                    // Replace with your custom field
                    updateIncident["new_resolutiontimeinminutes"] = resolutionMinutes;
                    updateIncident["new_netresolutiontimeinminutes"] = netResolutionMinutes;

                    service.Update(updateIncident);
                }
            }
        }
    }

    // Helper function to calculate business minutes (excluding weekends)

    private int GetResolutionMinutes(DateTime startUtc, DateTime endUtc, ITracingService tracingService)
    {
        if (endUtc <= startUtc) return 0;

        DateTime start = startUtc;
        DateTime end = endUtc;

        int totalMinutes = 0;
        DateTime currentDay = start.Date;

        while (currentDay <= end.Date)
        {
            if (currentDay.DayOfWeek != DayOfWeek.Saturday && currentDay.DayOfWeek != DayOfWeek.Sunday)
            {
                DateTime dayStart = currentDay == start.Date ? start : currentDay;
                DateTime dayEnd = currentDay == end.Date ? end : currentDay.AddDays(1);

                totalMinutes += (int)(dayEnd - dayStart).TotalMinutes;
            }

            currentDay = currentDay.AddDays(1);
        }
        return totalMinutes;
    }

    private int GetNetResolutionMinutes(DateTime startUtc, DateTime endUtc, ITracingService tracingService)
    {
        if (endUtc <= startUtc) return 0;

        DateTime start = startUtc;
        DateTime end = endUtc;

        int totalMinutes = 0;
        DateTime currentDay = start.Date;

        while (currentDay <= end.Date)
        {
            if (currentDay.DayOfWeek != DayOfWeek.Friday && currentDay.DayOfWeek != DayOfWeek.Saturday && currentDay.DayOfWeek != DayOfWeek.Sunday)
            {
                DateTime dayStart = currentDay == start.Date ? start : currentDay;
                DateTime dayEnd = currentDay == end.Date ? end : currentDay.AddDays(1);

                totalMinutes += (int)(dayEnd - dayStart).TotalMinutes;
            }
            else if (currentDay.DayOfWeek == DayOfWeek.Friday)
            {
                DateTime dayStart = currentDay == start.Date ? start : currentDay;
                DateTime dayEnd = currentDay == end.Date ? end : new DateTime(currentDay.Year, currentDay.Month, currentDay.Day, 18, 0, 0); // 6 PM on Friday

                if (dayStart <= dayEnd)
                {
                    totalMinutes += (int)(dayEnd - dayStart).TotalMinutes; // 18 hours on Friday
                }
            }
            else if (currentDay.DayOfWeek == DayOfWeek.Sunday)
            {
                DateTime dayStart = currentDay == start.Date ? start : new DateTime(currentDay.Year, currentDay.Month, currentDay.Day, 20, 0, 0); // 8 PM on Sunday
                DateTime dayEnd = currentDay == end.Date ? end : currentDay.AddDays(1);

                if (dayStart <= dayEnd)
                {
                    totalMinutes += (int)(dayEnd - dayStart).TotalMinutes; // 4 hours on Subday
                }
            }

            currentDay = currentDay.AddDays(1);
        }
        return totalMinutes;
    }
}
