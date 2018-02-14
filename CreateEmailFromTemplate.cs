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
    public class CreateEmailFromTemplate : CodeActivity
    {
        #region Input/Output Parameters

        [Input("Record Url")]
        [RequiredArgument]
        public InArgument<string> RecordUrl { get; set; }

        [RequiredArgument]
        [Input("Email Template")]
        [ReferenceTarget("template")]
        public InArgument<EntityReference> EmailTemplate { get; set; }

        [RequiredArgument]
        [Input("Sender")]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> Sender { get; set; }

        [Output("Email Reference")]
        [ReferenceTarget("email")]
        public OutArgument<EntityReference> EmailMessage { get; set; }

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
                EntityReference refEntity = parser.ConvertToEntityReference(service, RecordUrl.Get<string>(executionContext));
                EntityReference refTemplate = EmailTemplate.Get<EntityReference>(executionContext);

                    
                Guid emailId = CreateTemplatedEmail(service, refEntity, refTemplate.Id);
                Entity email = service.Retrieve("email", emailId, new ColumnSet("description"));
                string description = email.Contains("description") ? email.GetAttributeValue<string>("description") : string.Empty;

                List<Entity> fromList = new List<Entity>();
                EntityReference systemUser = Sender.Get<EntityReference>(executionContext);
                Entity user = new Entity("activityparty");
                user["partyid"] = systemUser;
                fromList.Add(user);

                UpdateEmailMessage(service, emailId, refEntity, fromList);
                EmailMessage.Set(executionContext, new EntityReference("email", emailId));
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.CreateEmailFromTemplate: " + ex.Message);
            }
            catch (System.Exception ex)
            {
                throw new Exception("XrmWorkflowTools.CreateEmailFromTemplate: " + ex.Message);
            }
        }

        private Guid CreateTemplatedEmail(IOrganizationService service, EntityReference incident, Guid templateId)
        {
            InstantiateTemplateRequest request = new InstantiateTemplateRequest()
            {
                TemplateId = templateId,
                ObjectId = incident.Id,
                ObjectType = incident.LogicalName
            };

            try
            {
                InstantiateTemplateResponse response = (InstantiateTemplateResponse)service.Execute(request);
                
                Entity email = response.EntityCollection[0];
                Guid emailId = service.Create(email);
                return emailId;

            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.CreateEmailFromTemplate.CreateTemplatedEmail: " + ex.Message);
            }
            catch (System.Exception ex)
            {
                throw new Exception("XrmWorkflowTools.CreateEmailFromTemplate.CreateTemplatedEmail: " + ex.Message);
            }
        }

        private void SendEmailMessage(IOrganizationService service, Guid emailId)
        {
            SendEmailRequest request = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = string.Empty,
                IssueSend = true
            };

            SendEmailResponse response = (SendEmailResponse)service.Execute(request);
        }

        private void UpdateEmailMessage(IOrganizationService service, Guid emailId, EntityReference regardingObject, List<Entity> fromList)
        {
            Entity email = new Entity("email");
            email.Id = emailId;
            email["from"] = fromList.ToArray();
            email["regardingobjectid"] = regardingObject;

            try
            {
                service.Update(email);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("XrmWorkflowTools.CreateEmailFromTemplate.UpdateEmailMessage: " + ex.Message);
            }
            catch (System.Exception ex)
            {
                throw new Exception("XrmWorkflowTools.CreateEmailFromTemplate.UpdateEmailMessage: " + ex.Message);
            }
        }

        private string GetTemplateContent(IOrganizationService service, Guid templateId)
        {
            Entity template = service.Retrieve("template", templateId, new ColumnSet(true));
            string body = GetDataFromXml(template.Attributes["body"].ToString(), "match");
            return body;
        }

        private string GetDataFromXml(string value, string attributeName)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            XDocument document = XDocument.Parse(value);
            // get the Element with the attribute name specified  
            XElement element = document.Descendants().Where(ele => ele.Attributes().Any(attr => attr.Name == attributeName)).FirstOrDefault();
            return element == null ? string.Empty : element.Value;
        }

    }
}
