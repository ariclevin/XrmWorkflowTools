using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Xml.Linq;

namespace XrmWorkflowTools
{
    public class ChangeProcessStage : CodeActivity
    {
        #region Input/Output Parameters

        [Input("Record Url")]
        [RequiredArgument]
        public InArgument<string> RecordUrl { get; set; }

        [RequiredArgument]
        [Input("Process Name")]
        public InArgument<string> Process { get; set; }

        [RequiredArgument]
        [Input("Process Stage Name")]
        public InArgument<string> ProcessStage { get; set; }

        #endregion

        protected override void Execute(CodeActivityContext executionContext)
        {
            // Extract the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            try
            {
                //Create the context
                IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                DynamicUrlParser parser = new DynamicUrlParser();
                EntityReference currentRecord = parser.ConvertToEntityReference(service, RecordUrl.Get<string>(executionContext));

                Guid processId = RetrieveProcessId(service, Process.Get<string>(executionContext));
                if (processId != Guid.Empty)
                {
                    Guid stageId = RetrieveStageId(service, processId, ProcessStage.Get<string>(executionContext));
                    if (stageId != Guid.Empty)
                    {
                        UpdateStage(service, currentRecord, processId, stageId);
                    }
                }

            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.ChangeProcessStage.Execute: " + ex.Message);
            }
            catch (System.Exception ex)
            {
                throw new Exception("XrmWorkflowTools.ChangeProcessStage.Execute: " + ex.Message);
            }
        }

        private Guid RetrieveProcessId(IOrganizationService service, string processName)
        {
            Guid processId = Guid.Empty;
            QueryExpression query = new QueryExpression("workflow")
            {
                ColumnSet = new ColumnSet("workflowid"),
                Criteria =
                {
                    Conditions = 
                    {
                        // new ConditionExpression("statecode", ConditionOperator.Equal, 0),
                        new ConditionExpression("name", ConditionOperator.Equal, processName)
                    }
                }
            };
            try
            {
                EntityCollection results = service.RetrieveMultiple(query);
                if (results.Entities.Count > 0)
                {
                    processId = results.Entities[0].Id;
                }
                return processId;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.ChangeProcessStage.RetrieveProcessId: " + ex.Message);
            }
            catch (System.Exception ex)
            {
                throw new Exception("XrmWorkflowTools.ChangeProcessStage.RetrieveProcessId: " + ex.Message);
            }
        }

        private Guid RetrieveStageId(IOrganizationService service, Guid processId, string stageName)
        {
            Guid stageId = Guid.Empty;
            QueryExpression query = new QueryExpression("processstage")
            {
                ColumnSet = new ColumnSet("processstageid"),
                Criteria =
                {
                    Conditions = 
                    {
                        // new ConditionExpression("statecode", ConditionOperator.Equal, 0),
                        new ConditionExpression("stagename", ConditionOperator.Equal, stageName),
                        new ConditionExpression("processid", ConditionOperator.Equal, processId)
                    }
                }
            };
            try
            {
                EntityCollection results = service.RetrieveMultiple(query);
                if (results.Entities.Count > 0)
                {
                    stageId = results.Entities[0].Id;
                }
                return stageId;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.ChangeProcessStage.RetrieveStageId: " + ex.Message);
            }
            catch (System.Exception ex)
            {
                throw new Exception("XrmWorkflowTools.ChangeProcessStage.RetrieveStageId: " + ex.Message);
            }

        }

        private void UpdateStage(IOrganizationService service, EntityReference currentRecord, Guid processId, Guid stageId)
        {
            Entity current = new Entity(currentRecord.LogicalName);
            current.Id = currentRecord.Id;
            current["processid"] = processId;
            current["stageid"] = stageId;

            try
            {
                service.Update(current);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.ChangeProcessStage.UpdateStage: " + ex.Message);
            }
            catch (System.Exception ex)
            {
                throw new Exception("XrmWorkflowTools.ChangeProcessStage.UpdateStage: " + ex.Message);
            }

        }

    }
}
