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
    public class AssignChildRecords : CodeActivity
    {
        #region Input/Output Parameters

        [Input("Record Url")]
        [RequiredArgument]
        public InArgument<string> RecordUrl { get; set; }

        [Input("Relationship Name")]
        [RequiredArgument]
        public InArgument<string> RelationshipName { get; set; }

        [Input("User")]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> SystemUser { get; set; }

        [Input("Team")]
        [ReferenceTarget("team")]
        public InArgument<EntityReference> Team { get; set; }

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

                Guid ownerId = Guid.Empty;
                string ownerType = string.Empty;

                EntityReference owner = SystemUser.Get<EntityReference>(executionContext);
                if (owner != null)
                {
                    ownerId = owner.Id;
                    ownerType = owner.LogicalName;
                }
                else
                {
                    owner = Team.Get<EntityReference>(executionContext);
                    if (owner != null)
                    {
                        ownerId = owner.Id;
                        ownerType = owner.LogicalName;
                    }
                }

                if (ownerId != Guid.Empty)
                {
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
                                RetrieveAndUpdateRelatedRecords(service, primaryEntity, childEntityName, childEntityAttribute, ownerId, ownerType);
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
                            RetrieveAndUpdateRelatedRecords(service, primaryEntity, childEntityName, childEntityAttribute, ownerId, ownerType);
                        }
                    }
                }

            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.AssignChildRecords: " + ex.Message);
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

        private void RetrieveAndUpdateRelatedRecords(IOrganizationService service, EntityReference primaryEntity, string childEntityName, string childEntityAttribute, Guid ownerId, string ownerIdType)
        {
            if (childEntityName != null)
            {
                EntityCollection rc = RetrieveChildEntityRecords(service, childEntityName, childEntityAttribute, primaryEntity);
                if (rc.Entities.Count > 0)
                {
                    foreach (Entity entity in rc.Entities)
                    {
                        UpdateEntityOwner(service, childEntityName, entity.Id, ownerId, ownerIdType);
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
                throw new Exception("XrmWorkflowTools.AssignChildRecords.RetrieveChildEntityRecords: " + ex.Message);
            }
        }


        private void UpdateEntityOwner(IOrganizationService service, string entityName, Guid entityId, Guid ownerId, string ownerIdType)
        {
            AssignRequest request = new AssignRequest()
            {
                Assignee = new EntityReference(ownerIdType, ownerId),
                Target = new EntityReference(entityName, entityId)
            };

            try
            {
                AssignResponse response = (AssignResponse)service.Execute(request);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.AssignChildRecords.UpdateEntityOwner: " + ex.Message);
            }

        }
    }
}
