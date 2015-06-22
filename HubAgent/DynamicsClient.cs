using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// These namespaces are found in the Microsoft.Crm.Sdk.Proxy.dll assembly
// located in the SDK\bin folder of the SDK download.
using Microsoft.Crm.Sdk.Messages;

// These namespaces are found in the Microsoft.Xrm.Sdk.dll assembly
// located in the SDK\bin folder of the SDK download.
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

// These namespaces are found in the Microsoft.Xrm.Client.dll assembly
// located in the SDK\bin folder of the SDK download.
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;

namespace HubAgent
{
    public interface IDynamicsClient
    {
        void AssociateGovdeliveryEmail(string id, string govdId);
        void EnsureEmailGovdeliveryField();
        void EnsureEmailMetadata();
        void EnsureGovdeliveryPublisher();
        string LookupEmailByGovdeliveryId(string govdId);
        List<Email> RetrievePendingEmails(string from);
        void UpdateStatus(string id, string status);
    }

    public class DynamicsClient : IDynamicsClient
    {
        private Dictionary<string, int> emailState;
        private Dictionary<string, int> emailStatus;
        private IOrganizationService service;

        public DynamicsClient(string url, string username, string password)
        {
            emailStatus = new Dictionary<string, int>();
            emailState = new Dictionary<string, int>();
            service = new OrganizationService(CrmConnection.Parse("Url=" + url + "; Username=" + username + "; Password=" + password));
        }

        // Externally callabled methods

        /// <summary>
        /// Associate an email activity with a GovDelivery email by writing the email ID into a field on the entity
        /// </summary>
        /// <returns></returns>
        public void AssociateGovdeliveryEmail(string id, string govdId)
        {
            RetrieveRequest emailRequest = new RetrieveRequest()
            {
                ColumnSet = new ColumnSet("govd_id"),
                Target = new EntityReference("email", new Guid(id))
            };

            var email = retrieveEntity(emailRequest);
            email["govd_id"] = govdId;

            UpdateRequest updateEmail = new UpdateRequest()
            {
                Target = email
            };
            this.getService().Execute(updateEmail);
        }

        /// <summary>
        /// Ensures that a govd_id to hold GovDelivery email IDs exists on email activities
        /// </summary>
        /// <returns></returns>
        public void EnsureEmailGovdeliveryField()
        {
            AttributeMetadata[] emailMetadataAttributes = this.retrieveMetadataAttributes("email");
            if (!emailMetadataAttributes.Any(prop => prop.LogicalName.Equals("govd_id")))
            {
                CreateAttributeRequest createGovDeliveryRequest = new CreateAttributeRequest
                {
                    EntityName = "email",
                    Attribute = new StringAttributeMetadata()
                    {
                        SchemaName = "govd_id",
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                        MaxLength = 100,
                        FormatName = StringFormatName.Text,
                        DisplayName = new Label("GovDelivery Message Id", 1033),
                        Description = new Label("The GovDelivery Transactional Message ID for tracking purposes.", 1033)
                    }
                };
                this.getService().Execute(createGovDeliveryRequest);
            }
        }

        /// <summary>
        /// Ensures that the required statuses on email activities exist
        /// </summary>
        /// <returns></returns>
        public void EnsureEmailMetadata()
        {
            List<string> expectedStatus = new List<string>(){ "Draft", "Completed",  "Sent", "Received", "Canceled", "Pending Send", "Sending", "Failed" };
            foreach (OptionMetadata attribute in this.getOptionSet("email", "statuscode"))
            {
                emailStatus[attribute.Label.UserLocalizedLabel.Label] = (int)attribute.Value;
            }
            foreach (OptionMetadata attribute in this.getOptionSet("email", "statecode"))
            {
                emailState[attribute.Label.UserLocalizedLabel.Label] = (int)attribute.Value;
            }
            var remainderStatus = this.listSubtraction(expectedStatus, emailStatus.Keys.ToList());

            foreach (string status in remainderStatus)
            {
                InsertStatusValueRequest insertStatusValueRequest = new InsertStatusValueRequest()
                {
                    AttributeLogicalName = "statuscode",
                    EntityLogicalName = "email",
                    //1033 below represents localeId for the United States and English
                    Label = new Label(status, 1033),
                    StateCode = this.getState("Completed")
                };
                emailStatus[status] = this.insertStatus(insertStatusValueRequest);
            }
        }

        /// <summary>
        /// Ensure that a publisher is created so that we can prefix fields with govd
        /// </summary>
        /// <returns></returns>
        public void EnsureGovdeliveryPublisher()
        {
            Entity publisher = new Entity("publisher");
            publisher["customizationprefix"] = "govd";
            publisher["uniquename"] = "GovDelivery";
            publisher["friendlyname"] = "govdelivery";
            publisher["description"] = "GovDelivery publisher";

            QueryExpression query = new QueryExpression()
            {
                EntityName = "publisher",
                ColumnSet = new ColumnSet("customizationprefix"),
                Criteria = new FilterExpression()
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions =
                    {
                        new ConditionExpression()
                        {
                            AttributeName = "customizationprefix",
                            Operator = ConditionOperator.Equal,
                            Values = { "govd" }
                        }
                    }
                },

            };
            EntityCollection entities = this.getService().RetrieveMultiple(query);

            if (entities.Entities.Count == 0)
            {
                CreateRequest publisherRequest = new CreateRequest
                {
                    Target = publisher
                };
                this.getService().Execute(publisherRequest);
            }
        }

        /// <summary>
        /// Grab the first email activity record with a specific govd_id
        /// </summary>
        /// <returns></returns>
        public string LookupEmailByGovdeliveryId(string govdId)
        {
            QueryExpression query = new QueryExpression()
            {
                EntityName = "email",
                ColumnSet = new ColumnSet("govd_id"),
                Criteria = new FilterExpression()
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions =
                    {
                        new ConditionExpression()
                        {
                            AttributeName = "govd_id",
                            Operator = ConditionOperator.Equal,
                            Values = { govdId }
                        },
                    }
                },
            };

            EntityCollection entities = this.getService().RetrieveMultiple(query);
            if (entities.Entities.Count != 0)
            {
                return entities.Entities[0].Id.ToString();
            }
            return null;
        }

        /// <summary>
        /// Retrieve email activities sent from the address specified in the argument that are in "Pending Send" state
        /// </summary>
        /// <returns></returns>
        public List<Email> RetrievePendingEmails(string from)
        {
            List<Email> emails = new List<Email>();
            QueryExpression query = new QueryExpression()
            {
                EntityName = "email",
                ColumnSet = new ColumnSet("to", "subject", "description", "ownerid", "statuscode", "statecode", "createdon"),
                LinkEntities = {
                    new LinkEntity()
                    {
                        LinkFromEntityName = "email",
                        LinkFromAttributeName = "ownerid",
                        LinkToEntityName = "systemuser",
                        LinkToAttributeName = "systemuserid",
                        EntityAlias = "owner",
                        JoinOperator = JoinOperator.Inner,
                        Columns = new ColumnSet("internalemailaddress")
                    }
                },
                Criteria = new FilterExpression()
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions =
                    {
                        new ConditionExpression()
                        {
                            AttributeName = "statuscode",
                            Operator = ConditionOperator.Equal,
                            Values = { (int)this.getStatus("Pending Send") }
                        },
                        new ConditionExpression()
                        {
                            AttributeName = "sender",
                            Operator = ConditionOperator.Equal,
                            Values = { from }
                        }
                    }
                },
            };

            EntityCollection emailEntities = this.getService().RetrieveMultiple(query);

            Console.WriteLine("Retrieved {0} emails", emailEntities.Entities.Count);
            foreach (Entity email in emailEntities.Entities)
            {
                var sender = (AliasedValue)email["owner.internalemailaddress"];
                var payload = new Email()
                {
                    id = email.Id.ToString(),
                    from = sender.Value.ToString(),
                    to = parseRecipients(email, "to"),
                    subject = email["subject"].ToString(),
                    body = email["description"].ToString()
                };
                emails.Add(payload);
            }
            return emails;
        }

        /// <summary>
        /// Update the status of an email activity
        /// </summary>
        /// <returns></returns>
        public void UpdateStatus(string id, string status)
        {
            SetStateRequest stateRequest = new SetStateRequest()
            {
                EntityMoniker = new EntityReference
                {
                    Id = new Guid(id),
                    LogicalName = "email"
                },
                State = new OptionSetValue((int)this.getState("Completed")),
                Status = new OptionSetValue((int)this.getStatus(status))
            };
            this.getService().Execute(stateRequest);
        }

        // Utility Methods

        public virtual List<OptionMetadata> getOptionSet(string entityName, string fieldName)
        {

            var attReq = new RetrieveAttributeRequest();
            attReq.EntityLogicalName = entityName;
            attReq.LogicalName = fieldName;
            attReq.RetrieveAsIfPublished = true;

            var attResponse = (RetrieveAttributeResponse)this.getService().Execute(attReq);
            var attMetadata = (EnumAttributeMetadata)attResponse.AttributeMetadata;

            return attMetadata.OptionSet.Options.ToList();
        }

        public virtual IOrganizationService getService()
        {
            return service;
        }

        public virtual int getState(string state)
        {
            return emailState[state];
        }

        public virtual int getStatus(string status)
        {
            return emailStatus[status];
        }

        public virtual int insertStatus(InsertStatusValueRequest request)
        {
            return ((InsertStatusValueResponse)this.getService().Execute(request)).NewOptionValue;
        }

        public List<string> listSubtraction(List<string> first, List<string> second)
        {
            List<string> remainder = new List<string>();
            foreach (string element in first)
            {
                if (!second.Contains(element))
                {
                    remainder.Add(element);
                }
            }
            return remainder;
        }

        public virtual List<string> parseRecipients(Entity email, string field)
        {
            List<string> recipients = new List<string>();
            var entities = (EntityCollection)email[field];
            foreach (Entity entity in entities.Entities)
            {
                recipients.Add(entity["addressused"].ToString());
            }
            return recipients;
        }

        public virtual Entity retrieveEntity(RetrieveRequest request)
        {
            RetrieveResponse response = (RetrieveResponse)service.Execute(request);
            return response.Entity;
        }

        public virtual AttributeMetadata[] retrieveMetadataAttributes(string entity)
        {
            RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.All,
                LogicalName = entity
            };

            RetrieveEntityResponse response = ((RetrieveEntityResponse)service.Execute(retrieveEntityRequest));
            return response.EntityMetadata.Attributes;
        }
    }
}