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
    public class GenerateCompositeAddress : CodeActivity
    {
        #region Input/Output Parameters

        [Input("Record Url")]
        [RequiredArgument]
        public InArgument<string> RecordUrl { get; set; }

        [Input("Address Line 1")]
        [RequiredArgument]
        public InArgument<string> Street1 { get; set; }

        [Input("Address Line 2")]
        [RequiredArgument]
        public InArgument<string> Street2 { get; set; }

        [Input("Address City")]
        [RequiredArgument]
        public InArgument<string> City { get; set; }

        [Input("Address State/Province")]
        [RequiredArgument]
        public InArgument<string> StateOrProvince { get; set; }

        [Input("Address Zip/Postal Code")]
        [RequiredArgument]
        public InArgument<string> ZipPostalCode { get; set; }

        [Input("Address Country")]
        public InArgument<string> Country { get; set; }

        [Output("Full Address")]
        public OutArgument<string> FullAddress { get; set; }

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

                // Don't Really Care about which entity this is
                DynamicUrlParser parser = new DynamicUrlParser();
                EntityReference primaryEntity = parser.ConvertToEntityReference(service, RecordUrl.Get<string>(executionContext));

                string addressLine1 = Street1.Get<string>(executionContext);
                string addressLine2 = Street2.Get<string>(executionContext);
                string city = City.Get<string>(executionContext);
                string stateOrProvince = StateOrProvince.Get<string>(executionContext);
                string zipCode = ZipPostalCode.Get<string>(executionContext);
                string countryName = Country.Get<string>(executionContext);

                string rc = GenerateAddress(addressLine1, addressLine2, city, stateOrProvince, zipCode, countryName);
                FullAddress.Set(executionContext, rc);


            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("SBS.Workflow.SetStateChildRecords: " + ex.Message);
            }
        }

        private string GenerateAddress(string line1, string line2, string city, string stateOrProvince, string zipPostalCode, string country)
        {
            #region Address Format
            // Street 1
            // Street 2 (Optional)
            // City, State/Province Postal
            // Country
            #endregion

            StringBuilder sb = new StringBuilder();

            if (!string.IsNullOrEmpty(line1))
            {
                sb.Append(line1);
                sb.AppendLine();
            }

            if (!string.IsNullOrEmpty(line2))
            {
                sb.Append(line2);
                sb.AppendLine();
            }

            if (!string.IsNullOrEmpty(city))
            {
                sb.Append(city);
                if (!string.IsNullOrEmpty(stateOrProvince))
                {
                    sb.Append(", " + stateOrProvince);
                    if (!string.IsNullOrEmpty(zipPostalCode))
                    {
                        sb.Append(" " + zipPostalCode);
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(zipPostalCode))
                    {
                        sb.Append(" " + zipPostalCode);
                    }
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(stateOrProvince))
                {
                    sb.Append(", " + stateOrProvince);
                    if (!string.IsNullOrEmpty(zipPostalCode))
                    {
                        sb.Append(" " + zipPostalCode);
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(zipPostalCode))
                    {
                        sb.Append(" " + zipPostalCode);
                    }
                }
            }

            if (!string.IsNullOrEmpty(country))
            {
                sb.AppendLine(country);
            }

            return sb.ToString();
        }
    }
}
