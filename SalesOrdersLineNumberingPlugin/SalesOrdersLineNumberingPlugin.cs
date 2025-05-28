using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Runtime.Remoting.Services;
using System.ServiceModel;

namespace SalesOrdersLineNumberingPlugin
{
    public class SalesOrdersLineNumberingPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("Starting plugin...");

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity entity)
            {

                try
                {
                    if (entity.LogicalName != "salesorderdetail") return;

                    if (!entity.Contains("salesorderid")) return;
                    var orderRef = (EntityReference)entity["salesorderid"];

                    // Query existing salesorderdetails for this order
                    var query = new QueryExpression("salesorderdetail")
                    {
                        ColumnSet = new ColumnSet("omnisync_lineitemnumber"),
                        Criteria =
                        {
                            Conditions =
                            {
                                new ConditionExpression("salesorderid", ConditionOperator.Equal, orderRef.Id)
                            }
                        }
                    };

                    var existingDetails = service.RetrieveMultiple(query).Entities;

                    tracingService.Trace("Before the calculation");

                    int nextLineNumber = 1;
                    if (existingDetails.Count > 0)
                    {
                        var maxLine = existingDetails
                            .Where(e => e.Attributes.Contains("omnisync_lineitemnumber"))
                            .Select(e => (int)e["omnisync_lineitemnumber"])
                            .DefaultIfEmpty(0)
                            .Max();

                        nextLineNumber = maxLine + 1;
                    }

                    // Set the new line number
                    entity["omnisync_lineitemnumber"] = nextLineNumber;

                    tracingService.Trace("Next line number " +  nextLineNumber);
                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in SalesOrdersLineNumberingPlugin.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("SalesOrdersLineNumberingPlugin: {0}", ex.ToString());
                    throw;
                }

                
            }
        }
    }
}
