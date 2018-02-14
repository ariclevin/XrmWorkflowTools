using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;

namespace SBS.Workflow
{
    /// <summary> 
    /// Used to parse the Dynamics CRM 'Record Url (Dynamic)' that can be created by workflows and dialogs 
    /// </summary> 
    public class DynamicUrlParser
    {
        public EntityReference ConvertToEntityReference(IOrganizationService service, string recordReference)
        {
            Uri uriResult;

            if (Uri.TryCreate(recordReference, UriKind.Absolute, out uriResult))
            {
                return ParseUrlToEntityReference(service, recordReference);
            }

            try
            {
                var jsonEntityReference = JsonConvert.DeserializeObject<JsonEntityReference>(recordReference);

                return new EntityReference(jsonEntityReference.LogicalName, jsonEntityReference.Id);
            }
            catch (Exception e)
            {
                throw new Exception("Error converting string '{recordReference}' to EntityReference - {e.Message}", e);
            }
        }

        private EntityReference ParseUrlToEntityReference(IOrganizationService service, string url)
        {
            Uri uri = new Uri(url);

            int found = 0;
            int entityTypeCode = 0;
            Guid id = Guid.Empty;

            var parameters = uri.Query.TrimStart('?').Split('&');
            foreach (var param in parameters)
            {
                var nameValue = param.Split('=');
                switch (nameValue[0])
                {
                    case "etc":
                        entityTypeCode = int.Parse(nameValue[1]);
                        found++;
                        break;
                    case "id":
                        id = new Guid(nameValue[1]);
                        found++;
                        break;
                }
                if (found > 1) break;
            }

            if (id == Guid.Empty)
                return null;

            MetadataFilterExpression entityFilter = new MetadataFilterExpression(LogicalOperator.And);
            entityFilter.Conditions.Add(new MetadataConditionExpression("ObjectTypeCode ", MetadataConditionOperator.Equals, entityTypeCode));
            MetadataPropertiesExpression propertyExpression = new MetadataPropertiesExpression { AllProperties = false };
            propertyExpression.PropertyNames.Add("LogicalName");
            EntityQueryExpression entityQueryExpression = new EntityQueryExpression()
            {
                Criteria = entityFilter,
                Properties = propertyExpression
            };

            RetrieveMetadataChangesRequest retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest()
            {
                Query = entityQueryExpression
            };

            RetrieveMetadataChangesResponse response = (RetrieveMetadataChangesResponse)service.Execute(retrieveMetadataChangesRequest);

            if (response.EntityMetadata.Count >= 1)
            {
                return new EntityReference(response.EntityMetadata[0].LogicalName, id);
            }

            return null;
        }
    }
}