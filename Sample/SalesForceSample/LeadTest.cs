using Pagos.Designer.Interfaces.External.CustomHooks;
using Pagos.Designer.Interfaces.External.Messaging;
using Pagos.SpreadsheetWeb.Web.Api.Objects.Calculation;
using SalesForceSample.sforce;
using System;
using System.Linq;
using System.Net;
using System.ServiceModel;

namespace SalesForceSample
{

    /// <summary>
    /// The custom code class, implementing relevant interfaces for desired 
    /// actions (i.e. AfterCalculation in this scenario).
    /// </summary>
    public class LeadTest : IAfterCalculation
    {
        private string accountName = "your_salesforce_account_email";
        private string accountPwd = "password";

        // example of using the constructor of a class to set security protocol options
        public LeadTest()
        {
            ServicePointManager.SecurityProtocol =
                SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
        }

        /// <summary>
        /// After calculation occurs in the calculation engine, provide a hook
        /// to perform additional custom actions.
        /// </summary>
        /// <param name="request">The request that was sent to the calculation engine.</param>
        /// <param name="response">The response that came back from the calculation engine.</param>
        /// <returns></returns>
        public ActionableResponse AfterCalculation(CalculationRequest request, CalculationResponse response)
        {
            // get values from named ranges of Excel file
            var name = request.Inputs.FirstOrDefault(x => x.Ref == "iName");
            var mail = request.Inputs.FirstOrDefault(x => x.Ref == "iEmail");

            // prepare the SOAP client
            BasicHttpBinding b = new BasicHttpBinding();
            EndpointAddress ea = new EndpointAddress("https://login.salesforce.com/services/Soap/c/43.0");
            b.Name = "Soap";
            b.Namespace = "sforce.Soap";
            b.Security.Mode = BasicHttpSecurityMode.Transport;

            SoapClient client = new SoapClient(b, ea);

            SessionHeader header;
            SoapClient apiClient;
            LoginResult loginResult;

            ActionableResponse ar = new ActionableResponse();
            try
            {
                // authenticate to a Salesforce Web Service
                loginResult = client.login(null, accountName, accountPwd);
                header = new sforce.SessionHeader();
                header.sessionId = loginResult.sessionId;
                EndpointAddress eaApiClient = new EndpointAddress(loginResult.serverUrl);

                // initialize API client
                apiClient = new SoapClient(b, eaApiClient);

                // create a Lead data
                Lead lead = new Lead()
                {
                    // use same value for FirstName and LastName
                    FirstName = name.Value[0][0].Value,
                    LastName = name.Value[0][0].Value,
                    Email = mail.Value[0][0].Value,
                    Company = "Company Name",
                    LeadSource = "Advertisement"
                };

                AssignmentRuleHeader arh = new AssignmentRuleHeader();
                QueryResult qr = null;

                LimitInfo[] li;
                SaveResult[] sr;

                // make an API request and add a Lead record to the Salesforce database
                apiClient.create(
                    header, // sessionheader
                    arh, // assignmentruleheader
                    null, // mruheader
                    null, // allowfieldtrunctionheader
                    null, // disablefeedtrackingheader
                    null, // streamingenabledheader
                    null, // allornoneheader
                    null, // duplicateruleheader
                    null, // localeoptions
                    null, // debuggingheader
                    null, // packageversionheader
                    null, // emailheader
                    new sObject[] { lead },
                    out li,
                    out sr);

                foreach (SaveResult s in sr)
                {
                    if (s.success)
                    {
                        ar.Success = true;
                        // for demonstration purpose we can set the id of added Lead record to a control visible on the page
                        response.Outputs.FirstOrDefault(x => x.Ref == "Response").Value[0][0].Value = s.id;
                    }
                    else
                    {
                        // set error messages to send back to a page
                        foreach (var e in s.errors)
                            ar.Messages.Add(e.message);

                        ar.Success = false;
                    }
                }
            }
            catch (Exception ex)
            {
                // set error message to send back to a page
                ar.Messages.Add(ex.Message);
                ar.Success = false;
            }

            return ar;
        }
    }
}