using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Email_Send
{
    class Program
    {
        static void Main(string[] args)
        {
            string authType = "OAuth";
            string userName = "shahzaib@SHAHZAIBSAFDAR1.onmicrosoft.com";
            string password = "safdar786ALI!";
            string url = "https://org666f01ac.crm.dynamics.com";
            string appId = "51f81489-12ee-4a9e-aaae-a2591f45987d";
            string reDirectURI = "app://58145B91-0C36-4500-8554-080854F2AC97";
            string loginPrompt = "Auto";
            string ConnectionString = string.Format("AuthType = {0};Username = {1};Password = {2}; Url = {3}; AppId={4}; RedirectUri={5};LoginPrompt={6}", authType, userName, password, url, appId, reDirectURI, loginPrompt);



            CrmServiceClient svc = new CrmServiceClient(ConnectionString);

            if (svc.IsReady)
            {
                Console.WriteLine("CONNECTION IS OKAY");
                // Get a system user to send the email (From: field)
                WhoAmIRequest systemUserRequest = new WhoAmIRequest();
                WhoAmIResponse systemUserResponse = (WhoAmIResponse)svc.Execute(systemUserRequest);
                Guid userId = systemUserResponse.UserId;

                // We will get the contact id for SHAHZAIB using Retrieve
                ConditionExpression conditionExpression = new ConditionExpression();
                conditionExpression.AttributeName = "firstname";
                conditionExpression.Operator = ConditionOperator.Equal;
                conditionExpression.Values.Add("shahzaib");

                FilterExpression filterExpression = new FilterExpression();
                filterExpression.Conditions.Add(conditionExpression);

                QueryExpression query = new QueryExpression("contact");
                query.ColumnSet.AddColumns("contactid");
                query.Criteria.AddFilter(filterExpression);

                EntityCollection entityCollection = svc.RetrieveMultiple(query);
                foreach (var a in entityCollection.Entities)
                {
                    // We will send the email from this current user
                    Entity fromActivityParty = new Entity("activityparty");
                    Entity toActivityParty = new Entity("activityparty");

                    Guid contactId = (Guid)a.Attributes["contactid"];

                    fromActivityParty["partyid"] = new EntityReference("systemuser", userId);
                    toActivityParty["partyid"] = new EntityReference("contact", contactId);

                    Entity email = new Entity("email");
                    email["from"] = new Entity[] { fromActivityParty };
                    email["to"] = new Entity[] { toActivityParty };
                    email["regardingobjectid"] = new EntityReference("contact", contactId);
                    email["subject"] = "This is the subject";
                    email["description"] = "This is the description.";
                    email["directioncode"] = true;
                    Guid emailId = svc.Create(email);

                    // Use the SendEmail message to send an e-mail message.
                    SendEmailRequest sendEmailRequest = new SendEmailRequest
                    {
                        EmailId = emailId,
                        TrackingToken = "",
                        IssueSend = true
                    };

                    SendEmailResponse sendEmailresp = (SendEmailResponse)svc.Execute(sendEmailRequest);
                    Console.WriteLine("Email sent");
                    Console.ReadLine();



                }



                    #region RELATED ENTITY
                    //              string myemail = "ranashahzaibulhassan@gamil.com";
                    //              var fetchXml = @"< fetch distinct = 'false' mapping = 'logical' output - format = 'xml - platform' version = '1.0' >
                    //< entity name = 'lead' >

                    //       < attribute name = 'fullname' />
                    //< attribute name = 'companyname' />
                    // < attribute name = 'telephone1' />
                    //< attribute name = 'leadid' />
                    //< order descending = 'false' attribute = 'fullname' /> < link - entity name = 'contact' alias = 'ak' link - type = 'inner' to = 'customerid' from = 'contactid' >
                    //< filter type = 'and' > < condition attribute = 'parentcustomerid' value = '{9BAC5E21-5D36-EC11-B6E6-000D3A5B2F35}' uitype = 'contact' uiname = 'li op' operator= 'eq' />
                    // </ filter > </ link - entity ></ entity > </ fetch > ";
                    //              //  fetchXml = String.Format(fetchXml, myemail);
                    //              EntityCollection contacts = svc.RetrieveMultiple(new FetchExpression(fetchXml));
                    //              Console.WriteLine("Total record: " + contacts.Entities.Count);
                    //              foreach (var contact in contacts.Entities)
                    //              {
                    //                  Console.WriteLine("Id" + contact.Id);
                    //                  if (contact.Contains("fullname") && contact["fullname"] != null)
                    //                  {
                    //                      Console.WriteLine("Contact Name: " + contact["fullname"]);
                    //                  }
                    //              }
                    //              Console.ReadLine();
                    //          }
                    #endregion
                }

            else
            {
                Console.WriteLine("Failed to Established Connection!!!");
            }
            Console.ReadLine();
        }
    }
}
