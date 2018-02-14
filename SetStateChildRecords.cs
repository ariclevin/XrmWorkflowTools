using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Crm.Sdk.Messages;

namespace XrmWorkflowTools
{
    public class SetStateChildRecords : CodeActivity
    {
        #region Input/Output Parameters

        [Input("Record Url")]
        [RequiredArgument]
        public InArgument<string> RecordUrl { get; set; }

        [Input("Relationship Name")]
        [RequiredArgument]
        public InArgument<string> RelationshipName { get; set; }

        [Input("State Code")]
        [RequiredArgument]
        public InArgument<int> StateCode { get; set; }

        [Input("Status Code")]
        [RequiredArgument]
        public InArgument<int> StatusCode { get; set; }

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
                EntityReference primaryEntity = parser.ConvertToEntityReference(service, RecordUrl.Get<string>(executionContext));
                string relationshipName = RelationshipName.Get<string>(executionContext);

                string childEntityName = null, childEntityAttribute = null;
                if (relationshipName.Contains(';'))
                {
                    string[] relationshipNames = relationshipName.Split(';');
                    foreach (string rel in relationshipNames)
                    {
                        OneToManyRelationshipMetadata oneToNRelationship = RetrieveRelationshipInfo(service, primaryEntity, relationshipName);
                        if (oneToNRelationship != null)
                        {
                            childEntityName = oneToNRelationship.ReferencingEntity;
                            childEntityAttribute = oneToNRelationship.ReferencingAttribute;
                            RetrieveAndUpdateRelatedRecords(service, primaryEntity, childEntityName, childEntityAttribute, StateCode.Get<int>(executionContext), StatusCode.Get<int>(executionContext));
                        }
                    }
                }
                else
                {
                    OneToManyRelationshipMetadata oneToNRelationship = RetrieveRelationshipInfo(service, primaryEntity, relationshipName);
                    if (oneToNRelationship != null)
                    {
                        childEntityName = oneToNRelationship.ReferencingEntity;
                        childEntityAttribute = oneToNRelationship.ReferencingAttribute;
                        RetrieveAndUpdateRelatedRecords(service, primaryEntity, childEntityName, childEntityAttribute, StateCode.Get<int>(executionContext), StatusCode.Get<int>(executionContext));
                    }
                }



            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.SetStateChildRecords: " + ex.Message);
            }
        }

        private OneToManyRelationshipMetadata RetrieveRelationshipInfo(IOrganizationService service, EntityReference primaryEntity, string relationshipName)
        {
            Relationship relationship = new Relationship(relationshipName);
            RetrieveEntityRequest request = new RetrieveEntityRequest()
            {
                LogicalName = primaryEntity.LogicalName,
                EntityFilters = EntityFilters.Relationships,
                RetrieveAsIfPublished = true
            };
            RetrieveEntityResponse response = (RetrieveEntityResponse)service.Execute(request);

            OneToManyRelationshipMetadata oneToNRelationship = response.EntityMetadata.OneToManyRelationships.FirstOrDefault(r => r.SchemaName == relationshipName);
            return oneToNRelationship;
        }

        private void RetrieveAndUpdateRelatedRecords(IOrganizationService service, EntityReference primaryEntity, string childEntityName, string childEntityAttribute, int stateCode, int statusCode)
        {
            if (childEntityName != null)
            {
                EntityCollection rc = RetrieveChildEntityRecords(service, childEntityName, childEntityAttribute, primaryEntity);
                if (rc.Entities.Count > 0)
                {
                    foreach (Entity entity in rc.Entities)
                    {
                        UpdateEntityStatus(service, childEntityName, entity.Id, stateCode, statusCode);
                    }
                }
            }
        }

        private EntityCollection RetrieveChildEntityRecords(IOrganizationService service, string entityName, string attributeName, EntityReference primaryEntity)
        {
            QueryExpression query = new QueryExpression(entityName)
            {
                ColumnSet = new ColumnSet(entityName + "id"),
                Criteria = new FilterExpression(LogicalOperator.And)
                {
                    Conditions =
                    {
                        new ConditionExpression(attributeName, ConditionOperator.Equal, primaryEntity.Id)
                    }
                }
            };

            try
            {
                EntityCollection results = service.RetrieveMultiple(query);
                return results;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.SetStateChildRecords.RetrieveChildEntityRecords: " + ex.Message);
            }
        }


        private void UpdateEntityStatus(IOrganizationService service, string entityName, Guid entityId, int stateCode, int statusCode)
        {
            EntityReference moniker = new EntityReference(entityName, entityId);

            SetStateRequest request = new SetStateRequest()
            {
                EntityMoniker = moniker,
                State = new OptionSetValue(stateCode),
                Status = new OptionSetValue(statusCode)
            };

            try
            {
                SetStateResponse response = (SetStateResponse)service.Execute(request);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.SetStateChildRecords.RetrieveChildEntityRecords: " + ex.Message);
            }

        }
    }
}
