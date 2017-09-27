using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Dynamics.CRM.YK.Workflows
{
    /// <summary>
    /// This custom workflow assembly returns the total active process session count for particular record for particular workflow.
    /// </summary>
    public class GetActiveProcessSessionCount : CodeActivity
    {
        /// <summary>
        /// This is the output parameter of the workflow assembly where total session count will be set.
        /// </summary>
        [Output("Active Process Count")]
        public OutArgument<int> activeProcessCount { get; set; }

        /// <summary>
        /// This is an input parameter which contains lookup to workflow for which the session to be fetched.
        /// </summary>
        [Input("Workflow")]
        [ReferenceTarget("workflow")]
        public InArgument<EntityReference> workflow { get; set; }

        /// <summary>
        /// Execution starts here.
        /// </summary>
        /// <param name="executionContext">Execution context passed by CRM.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.PrimaryEntityId.Equals(Guid.Empty) || string.IsNullOrEmpty(context.PrimaryEntityName))
                {
                    tracingService.Trace("Either primary entity id or primary entity name is missing");
                    return;
                }
                tracingService.Trace("Primary Entity Name: " + context.PrimaryEntityName);
                tracingService.Trace("Primary Entity Id: " + context.PrimaryEntityId);
                EntityReference process = workflow.Get(executionContext);
                tracingService.Trace("Workflow Id: " + process.Id.ToString());
                Entity workflowRecord = service.Retrieve(process.LogicalName, process.Id, new ColumnSet("activeworkflowid"));
                if (workflowRecord.Contains("activeworkflowid"))
                {
                    tracingService.Trace("Active Workflow Id: " + workflowRecord.GetAttributeValue<EntityReference>("activeworkflowid").Id.ToString());
                    activeProcessCount.Set(executionContext, GetTotalActiveSessionCount(context.PrimaryEntityId, context.PrimaryEntityName, workflowRecord.GetAttributeValue<EntityReference>("activeworkflowid").Id, service, tracingService));
                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

        /// <summary>
        /// Fetches active process session counts.
        /// </summary>
        /// <param name="recordId">Guid of the record.</param>
        /// <param name="entityName">Logical name of the record's entity.</param>
        /// <param name="workflowId">Guid of the workflow.</param>
        /// <param name="service">Instance of IOrganizationService interface.</param>
        /// <param name="tracingService">Instance of ITracingService interface.</param>
        /// <returns>Total active process session counts.</returns>
        private int GetTotalActiveSessionCount(Guid recordId, string entityName, Guid workflowId, IOrganizationService service, ITracingService tracingService)
        {
            string queryXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='asyncoperation'>
                                    <attribute name='asyncoperationid' />
                                    <filter type='and'>
                                      <condition attribute='statuscode' operator='in'>
                                        <value>20</value>
                                        <value>10</value>
                                        <value>0</value>
                                      </condition>
                                      <condition attribute='operationtype' operator='eq' value='10' />
                                      <condition attribute='regardingobjectid' operator='eq' uitype='" + entityName + @"' value='" + recordId.ToString() + @"' />
                                      <condition attribute='owningextensionid' operator='eq' value='" + workflowId.ToString() + @"' />
                                    </filter>
                                  </entity>
                                </fetch>";
            EntityCollection sessions = service.RetrieveMultiple(new FetchExpression(queryXml));
            tracingService.Trace("Result count:" + sessions.Entities.Count);
            return sessions.Entities.Count;
        }
    }
}
