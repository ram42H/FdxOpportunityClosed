using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace FdxOpportunityClosed
{
    public class ValidateOpportunityClosedAsWon : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins....
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Obtain execution contest from the service provider....
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            int step = 0;
            Guid opportunityId = Guid.Empty;

            if (context.InputParameters.Contains("OpportunityClose") && context.InputParameters["OpportunityClose"] is Entity)
            {
                step = 1;
                Entity opportunityEntity = (Entity)context.InputParameters["OpportunityClose"];

                //if (opportunityEntity.LogicalName != "OpportunityClose")
                //    return;

                try
                {
                    step = 2;
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    //Get current user information....
                    step = 3;
                    WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

                    step = 4;
                    if (opportunityEntity.Attributes.Contains("opportunityid") && opportunityEntity.Attributes["opportunityid"] != null)
                    {
                        EntityReference entityRef = (EntityReference)opportunityEntity.Attributes["opportunityid"];
 
                        if (entityRef.LogicalName == "opportunity")
                        {
                            opportunityId = entityRef.Id;
                        }
                    } 

                    step = 4;
                    QueryExpression oppProductQuery = CRMQueryExpression.getQueryExpression("opportunityproduct", new ColumnSet("opportunityid", "opportunityproductid"), new CRMQueryExpression[] { new CRMQueryExpression("opportunityid", ConditionOperator.Equal, opportunityId) });
                    step = 5;
                    EntityCollection opportunityProduct = service.RetrieveMultiple(oppProductQuery);
                    step = 6;
                    if (!(opportunityProduct.Entities.Count > 0))
                        throw new InvalidPluginExecutionException(OperationStatus.Failed, "Please add products in the opportunity before closing it as won. \n\n\n");

                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException(string.Format("An error occurred in the ValidateOpportunityClosedAsWon plug-in at Step {0}.", step), ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("ValidateOpportunityClosedAsWon: step {0}, {1}", step, ex.ToString());
                    throw;
                }
            }
        }
    }
}
