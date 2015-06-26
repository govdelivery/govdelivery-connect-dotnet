using System;
using System.Collections.Generic;
using System.Reflection;

using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

using HubAgent;

namespace HubAgentTest
{
    [TestClass]
    public class DynamicsClientTests
    {
        [TestMethod]
        public void TestAssociateGovdeliveryEmail()
        {
            var service = new Mock<IOrganizationService>();
            service.Setup(dyn => dyn.Execute(It.IsAny<UpdateRequest>()));

            var client = new Mock<DynamicsClient>("http://blah.test.com", "someone", "password");
            client.Setup(obj => obj.retrieveEntity(It.IsAny<RetrieveRequest>())).Returns(new Entity());
            client.Setup(obj => obj.getService()).Returns(service.Object);

            client.Object.AssociateGovdeliveryEmail("5cd54c07-3391-4bc0-a68e-911c2a38ed0e", "2");
            service.Verify(dyn => dyn.Execute(It.IsAny<UpdateRequest>()));
            client.Verify(obj => obj.retrieveEntity(It.IsAny<RetrieveRequest>()));
            client.Verify(obj => obj.getService());
        }

        [TestMethod]
        public void TestEnsureEmailGovdeliveryFieldPreexistingField()
        {
            var service = new Mock<IOrganizationService>();
            service.Setup(dyn => dyn.Execute(It.IsAny<CreateAttributeRequest>()));

            var client = new Mock<DynamicsClient>("http://blah.test.com", "someone", "password");
            client.Setup(obj => obj.getService()).Returns(service.Object);

            var attributes = new AttributeMetadata[]
                             {
                                 new AttributeMetadata(){ LogicalName = "govd_id" }
                             };
            client.Setup(obj => obj.retrieveMetadataAttributes(It.IsAny<string>())).Returns(attributes);

            client.Object.EnsureEmailGovdeliveryField();
            client.Verify(obj => obj.retrieveMetadataAttributes(It.IsAny<string>()));
            service.Verify(dyn => dyn.Execute(It.IsAny<CreateAttributeRequest>()), Times.Never());
        }

        [TestMethod]
        public void TestEnsureEmailGovdeliveryFieldNoField()
        {
            var service = new Mock<IOrganizationService>();
            service.Setup(dyn => dyn.Execute(It.IsAny<CreateAttributeRequest>()));

            var client = new Mock<DynamicsClient>("http://blah.test.com", "someone", "password");
            client.Setup(obj => obj.getService()).Returns(service.Object);

           var attributes = new AttributeMetadata[]{};
            client.Setup(obj => obj.retrieveMetadataAttributes(It.IsAny<string>())).Returns(attributes);

            client.Object.EnsureEmailGovdeliveryField();
            client.Verify(obj => obj.retrieveMetadataAttributes(It.IsAny<string>()));
            service.Verify(dyn => dyn.Execute(It.IsAny<CreateAttributeRequest>()));
        }

        [TestMethod]
        public void TestEnsureEmailMetadata()
        {
            var statuses = new List<OptionMetadata>()
            {
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Draft"
                        }
                    },
                    Value = 1
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Completed"
                        }
                    },
                    Value = 2
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Sent"
                        }
                    },
                    Value = 3
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Received"
                        }
                    },
                    Value = 4
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Canceled"
                        }
                    },
                    Value = 5
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Pending Send"
                        }
                    },
                    Value = 6
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Sending"
                        }
                    },
                    Value = 7
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Failed"
                        }
                    },
                    Value = 8
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Bounced"
                        }
                    },
                    Value = 9
                }
            };
            var states = new List<OptionMetadata>()
            {
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Completed"
                        }
                    },
                    Value = 1
                }
            };

            var service = new Mock<IOrganizationService>();
            service.Setup(dyn => dyn.Execute(It.IsAny<InsertStatusValueRequest>())).Returns(new InsertStatusValueResponse());

            var client = new Mock<DynamicsClient>("http://blah.test.com", "someone", "password");
            client.Setup(obj => obj.getService()).Returns(service.Object);
            client.Setup(obj => obj.insertStatus(It.IsAny<InsertStatusValueRequest>())).Returns(1);
            client.Setup(obj => obj.getOptionSet(It.IsAny<string>(), "statuscode")).Returns(statuses);
            client.Setup(obj => obj.getOptionSet(It.IsAny<string>(), "statecode")).Returns(states);

            client.Object.EnsureEmailMetadata();
            client.Verify(obj => obj.insertStatus(It.IsAny<InsertStatusValueRequest>()), Times.Never);
        }

        [TestMethod]
        public void TestEnsureEmailMetadataMissingMetadata()
        {
            var statuses = new List<OptionMetadata>()
            {
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Draft"
                        }
                    },
                    Value = 1
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Pending Send"
                        }
                    },
                    Value = 6
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Sending"
                        }
                    },
                    Value = 7
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Failed"
                        }
                    },
                    Value = 8
                },
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Bounced"
                        }
                    },
                    Value = 9
                },
            };
            var states = new List<OptionMetadata>()
            {
                new OptionMetadata
                {
                    Label = new Label()
                    {
                        UserLocalizedLabel = new LocalizedLabel()
                        {
                            Label = "Completed"
                        }
                    },
                    Value = 1
                }
            };

            var service = new Mock<IOrganizationService>();
            service.Setup(dyn => dyn.Execute(It.IsAny<InsertStatusValueRequest>())).Returns(new InsertStatusValueResponse());

            var client = new Mock<DynamicsClient>("http://blah.test.com", "someone", "password");
            client.Setup(obj => obj.getService()).Returns(service.Object);
            client.Setup(obj => obj.insertStatus(It.IsAny<InsertStatusValueRequest>())).Returns(1);
            client.Setup(obj => obj.getOptionSet(It.IsAny<string>(), "statuscode")).Returns(statuses);
            client.Setup(obj => obj.getOptionSet(It.IsAny<string>(), "statecode")).Returns(states);

            client.Object.EnsureEmailMetadata();
            client.Verify(obj => obj.insertStatus(It.IsAny<InsertStatusValueRequest>()), Times.Exactly(4));
        }

        [TestMethod]
        public void TestEnsureGovdeliveryPublisherPreexistingPublisher()
        {
            var entities = new EntityCollection()
            {
                Entities = { new Entity() }
            };

            var service = new Mock<IOrganizationService>();
            service.Setup(dyn => dyn.RetrieveMultiple(It.IsAny<QueryExpression>())).Returns(entities);
            service.Setup(dyn => dyn.Execute(It.IsAny<CreateRequest>()));

            var client = new Mock<DynamicsClient>("http://blah.test.com", "someone", "password");
            client.Setup(obj => obj.getService()).Returns(service.Object);

            client.Object.EnsureGovdeliveryPublisher();
            service.Verify(dyn => dyn.RetrieveMultiple(It.IsAny<QueryExpression>()));
            service.Verify(dyn => dyn.Execute(It.IsAny<CreateRequest>()), Times.Never);
        }

        [TestMethod]
        public void TestEnsureGovdeliveryPublisherNoPublisher()
        {
            var entities = new EntityCollection();

            var service = new Mock<IOrganizationService>();
            service.Setup(dyn => dyn.RetrieveMultiple(It.IsAny<QueryExpression>())).Returns(entities);
            service.Setup(dyn => dyn.Execute(It.IsAny<CreateRequest>()));

            var client = new Mock<DynamicsClient>("http://blah.test.com", "someone", "password");
            client.Setup(obj => obj.getService()).Returns(service.Object);

            client.Object.EnsureGovdeliveryPublisher();
            service.Verify(dyn => dyn.RetrieveMultiple(It.IsAny<QueryExpression>()));
            service.Verify(dyn => dyn.Execute(It.IsAny<CreateRequest>()));
        }

        [TestMethod]
        public void TestLookupEmailByGovdeliveryId()
        {
            var entities = new EntityCollection()
            {
                Entities = {
                    new Entity(){
                        Attributes = {
                            {"govd_id", "1"}
                        },
                        Id = new Guid("5cd54c07-3391-4bc0-a68e-911c2a38ed0e")
                    }
                }
            };

            var service = new Mock<IOrganizationService>();
            service.Setup(dyn => dyn.RetrieveMultiple(It.IsAny<QueryExpression>())).Returns(entities);

            var client = new Mock<DynamicsClient>("http://blah.test.com", "someone", "password");
            client.Setup(obj => obj.getService()).Returns(service.Object);

            var id = client.Object.LookupEmailByGovdeliveryId("1");
            service.Verify(dyn => dyn.RetrieveMultiple(It.IsAny<QueryExpression>()));

            Assert.AreEqual(id, "5cd54c07-3391-4bc0-a68e-911c2a38ed0e");
        }

        [TestMethod]
        public void TestLookupEmailByGovdeliveryIdNull()
        {
            var entities = new EntityCollection();

            var service = new Mock<IOrganizationService>();
            service.Setup(dyn => dyn.RetrieveMultiple(It.IsAny<QueryExpression>())).Returns(entities);

            var client = new Mock<DynamicsClient>("http://blah.test.com", "someone", "password");
            client.Setup(obj => obj.getService()).Returns(service.Object);

            var id = client.Object.LookupEmailByGovdeliveryId("1");
            service.Verify(dyn => dyn.RetrieveMultiple(It.IsAny<QueryExpression>()));

            Assert.IsNull(id);
        }

        [TestMethod]
        public void TestUpdateStatus()
        {
            var service = new Mock<IOrganizationService>();
            service.Setup(dyn => dyn.Execute(It.IsAny<SetStateRequest>()));

            var client = new Mock<DynamicsClient>("http://blah.test.com", "someone", "password");
            client.Setup(obj => obj.getService()).Returns(service.Object);
            client.Setup(obj => obj.getStatus("Sent")).Returns(3);

            client.Object.UpdateStatus("5cd54c07-3391-4bc0-a68e-911c2a38ed0e", "Sent");
            service.Verify(dyn => dyn.Execute(It.IsAny<SetStateRequest>()));
        }

        [TestMethod]
        public void TestRetrievePendingEmails()
        {
            var entities = new EntityCollection()
            {
                Entities = {
                    new Entity(){
                        Attributes = {
                            {"owner.internalemailaddress", new AliasedValue("blah", "blah", "other_test@test.com")},
                            {"subject", "Test Message"},
                            {"description", "Email body here"},
                            {"to", new EntityCollection(){Entities = { new Entity(){Attributes = {{"addressused", "a@b.com"}}}}}}
                        }
                    },
                    new Entity(){
                        Attributes = {
                            {"owner.internalemailaddress", new AliasedValue("blah", "blah", "other_test@test.com")},
                            {"subject", "Test Message 2"},
                            {"description", "Email body here"},
                            {"to", new EntityCollection(){Entities = { new Entity(){Attributes = {{"addressused", "b@b.com"}}}}}}
                        }
                    }

                }
            };

            var service = new Mock<IOrganizationService>();
            service.Setup(dyn => dyn.RetrieveMultiple(It.IsAny<QueryExpression>())).Returns(entities);

            var client = new Mock<DynamicsClient>("http://blah.test.com", "someone", "password");
            client.Setup(obj => obj.getService()).Returns(service.Object);
            client.Setup(obj => obj.getStatus("Pending Send")).Returns(3);

            var emails = client.Object.RetrievePendingEmails("me@test.com");
            service.Verify(dyn => dyn.RetrieveMultiple(It.IsAny<QueryExpression>()));

            Assert.AreEqual(emails.Count, 2);
        }
    }
}