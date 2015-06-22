using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.ServiceModel;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Microsoft.Xrm.Sdk.Metadata;
using Moq;

using HubAgent;

namespace HubAgentTest
{
    public class FakeHttpMessageHandler : HttpMessageHandler
    {
        private HttpResponseMessage response;

        public FakeHttpMessageHandler(HttpResponseMessage response)
        {
            this.response = response;
        }

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var responseTask = new TaskCompletionSource<HttpResponseMessage>();
            responseTask.SetResult(response);

            return responseTask.Task;
        }
    }

    [TestClass]
    public class HubClientTest
    {

        [TestMethod]
        public void TestRetrieveConnectorsWithDynamicsEmailConnector()
        {
            var fakeResponseHandler = new FakeHttpMessageHandler(connectorResponse("Connectors::DynamicsEmail"));
            var fakeHttpClient = new HttpClient(fakeResponseHandler);
            fakeHttpClient.BaseAddress = new Uri("https://blah.test.com");

            var mockClient = new Mock<IDynamicsClient>();
            var hubClient = new HubClient("https://blah.test.com", "blah", mockClient.Object);
            var hubConnection = hubClient.GetType().GetField("connection", BindingFlags.NonPublic | BindingFlags.Instance);
            hubConnection.SetValue(hubClient, fakeHttpClient);

            hubClient.RetrieveConnectors().Wait();

            var dynamicsValue = getPrivateValue(typeof(HubClient), hubClient, "dynamics");
            Assert.IsNotNull(dynamicsValue);
        }

        [TestMethod]
        public void TestRetrieveConnectorsWithOtherConnector()
        {
            var fakeResponseHandler = new FakeHttpMessageHandler(connectorResponse("Connectors::ThirdPartyEmail"));
            var fakeHttpClient = new HttpClient(fakeResponseHandler);
            fakeHttpClient.BaseAddress = new Uri("https://blah.test.com");

            var mockClient = new Mock<IDynamicsClient>();
            var hubClient = new HubClient("https://blah.test.com", "blah", mockClient.Object);
            var hubConnection = hubClient.GetType().GetField("connection", BindingFlags.NonPublic | BindingFlags.Instance);
            hubConnection.SetValue(hubClient, fakeHttpClient);

            hubClient.RetrieveConnectors().Wait();

            var dynamicsValue = getPrivateValue(typeof(HubClient), hubClient, "dynamics");
            Assert.IsNull(dynamicsValue);
        }

        [TestMethod]
        public void TestRetrieveConnectorsWithNoConnector()
        {
            string nullString = null;
            var fakeResponseHandler = new FakeHttpMessageHandler(connectorResponse(nullString));
            var fakeHttpClient = new HttpClient(fakeResponseHandler);
            fakeHttpClient.BaseAddress = new Uri("https://blah.test.com");

            var mockClient = new Mock<IDynamicsClient>();
            var hubClient = new HubClient("https://blah.test.com", "blah", mockClient.Object);
            var hubConnection = hubClient.GetType().GetField("connection", BindingFlags.NonPublic | BindingFlags.Instance);
            hubConnection.SetValue(hubClient, fakeHttpClient);

            hubClient.RetrieveConnectors().Wait();

            var dynamicsValue = getPrivateValue(typeof(HubClient), hubClient, "dynamics");
            Assert.IsNull(dynamicsValue);
        }

        [TestMethod]
        public void TestUpdateEmailStatuses()
        {
            Link link1 = new Link() {
                rel = "actions",
                href = "/connectors/123/actions"
            };

            Event email_sent_event = new Event()
            {
                name = "email_sent",
                external_id = "23a4",
                message_id = 8080
            };

            Event email_status_event = new Event()
            {
                name = "email_status",
                message_id = 8081,
                status = "sent"
            };

            Event invalid_email_status_event = new Event()
            {
                name = "email_status",
                message_id = 8082,
                status = "invalid"
            };

            Event invalid_event = new Event()
            {
                name = "blah_blah",
                external_id = "zzz",
                message_id = 8083
            };

            List<Event> events = new List<Event>();
            events.Add(email_sent_event);
            events.Add(email_status_event);
            events.Add(invalid_email_status_event);
            events.Add(invalid_event);

            var connector = new Connector();
            connector.links = new List<Link>();
            connector.links.Add(link1);

            var fakeResponseHandler = new FakeHttpMessageHandler(eventResponse(events));
            var fakeHttpClient = new HttpClient(fakeResponseHandler);
            fakeHttpClient.BaseAddress = new Uri("https://blah.test.com");

            var mockClient = new Mock<IDynamicsClient>();
            mockClient.Setup(dyn => dyn.AssociateGovdeliveryEmail(It.IsAny<string>(), It.IsAny<string>()));
            mockClient.Setup(dyn => dyn.UpdateStatus(It.IsAny<string>(), It.IsAny<string>()));
            mockClient.Setup(dyn => dyn.LookupEmailByGovdeliveryId(It.IsAny<string>())).Returns("5cd54c07-3391-4bc0-a68e-911c2a38ed0e");
            var hubClient = new HubClient("https://blah.test.com", "blah", mockClient.Object);

            var hubConnection = hubClient.GetType().GetField("connection", BindingFlags.NonPublic | BindingFlags.Instance);
            hubConnection.SetValue(hubClient, fakeHttpClient);
            var hubDynamics = hubClient.GetType().GetField("dynamics", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            hubDynamics.SetValue(hubClient, connector);

            hubClient.UpdateEmailStatuses().Wait();

            mockClient.Verify(dyn => dyn.AssociateGovdeliveryEmail("23a4", "8080"));
            mockClient.Verify(dyn => dyn.UpdateStatus("5cd54c07-3391-4bc0-a68e-911c2a38ed0e","Sent"));
        }

        [TestMethod]
        public void TestSendEmailsSuccess()
        {
            List<Email> emails = new List<Email>()
            { new Email()
                {
                    to = new List<string>() { "recipient@example.com" },
                    subject = "An email subject",
                    body = "Body content of the email",
                    id = "55a"
                }
            };

            var connector = new Connector();
            connector.links = new List<Link>()
            {
                new Link()
                {
                    rel = "actions",
                    href = "/connectors/123/actions"
                }
            };

            var fakeResponseHandler = new FakeHttpMessageHandler(new HttpResponseMessage(HttpStatusCode.OK));
            var fakeHttpClient = new HttpClient(fakeResponseHandler);
            fakeHttpClient.BaseAddress = new Uri("https://blah.test.com");

            var mockClient = new Mock<IDynamicsClient>();
            mockClient.Setup(dyn => dyn.UpdateStatus(It.IsAny<string>(), It.IsAny<string>()));
            var hubClient = new HubClient("https://prefix.domain.com", "blah", mockClient.Object);

            var hubConnection = hubClient.GetType().GetField("connection", BindingFlags.NonPublic | BindingFlags.Instance);
            hubConnection.SetValue(hubClient, fakeHttpClient);
            var hubDynamics = hubClient.GetType().GetField("dynamics", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            hubDynamics.SetValue(hubClient, connector);

            hubClient.SendEmails(emails).Wait();

            mockClient.Verify(dyn => dyn.UpdateStatus("55a", "Sending"));
        }

        [TestMethod]
        public void TestSendEmailsFailure()
        {
            List<Email> emails = new List<Email>()
            { new Email()
                {
                    to = new List<string>() { "recipient@example.com" },
                    subject = "An email subject",
                    body = "Body content of the email"
                }
            };

            var connector = new Connector();
            connector.links = new List<Link>()
            {
                new Link()
                {
                    rel = "actions",
                    href = "/connectors/123/actions"
                }
            };

            var fakeResponseHandler = new FakeHttpMessageHandler(new HttpResponseMessage(HttpStatusCode.NotFound));
            var fakeHttpClient = new HttpClient(fakeResponseHandler);
            fakeHttpClient.BaseAddress = new Uri("https://blah.test.com");

            var mockClient = new Mock<IDynamicsClient>();
            mockClient.Setup(dyn => dyn.UpdateStatus(It.IsAny<string>(), It.IsAny<string>()));
            var hubClient = new HubClient("https://prefix.domain.com", "blah", mockClient.Object);

            var hubConnection = hubClient.GetType().GetField("connection", BindingFlags.NonPublic | BindingFlags.Instance);
            hubConnection.SetValue(hubClient, fakeHttpClient);
            var hubDynamics = hubClient.GetType().GetField("dynamics", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            hubDynamics.SetValue(hubClient, connector);

            hubClient.SendEmails(emails).Wait();

            mockClient.Verify(dyn => dyn.UpdateStatus(It.IsAny<string>(), It.IsAny<string>()), Times.Never);
        }

        /// <summary>
        /// Create a an HttpResponseMessage that returns Connectors for use with mocked HttpClients
        /// </summary>
        /// <param name="types">An array of the type of connector to create</param>
        /// <returns></returns>
        private HttpResponseMessage connectorResponse(string[] types)
        {
            var httpResponse = new HttpResponseMessage(HttpStatusCode.OK);

            MemoryStream stream = new MemoryStream();

            DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
            settings.UseSimpleDictionaryFormat = true;
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(List<Connector>), settings);

            var connectorList = new List<Connector>();

            if (types != null)
            {
                for (int i = 0; i < types.Length; i++ )
                {
                    connectorList.Add(new Connector()
                        {
                            id = 1,
                            type = types[i],
                            options = new Dictionary<string, string>()
                        {
                            {"option", "one"}
                        },
                            links = new List<Link>()
                        {
                            new Link()
                            {
                                rel = "test",
                                href = "/actions"
                            }
                        }
                    });
                }
            }

            serializer.WriteObject(stream, connectorList);

            httpResponse.Content = new StringContent(Encoding.UTF8.GetString(stream.ToArray()), Encoding.UTF8, "application/json");
            return httpResponse;
        }

        /// <summary>
        /// CreateHttpResponseMessage that returns a List of Events for use with mocked HttpClients
        /// </summary>
        /// <param name="events">A List of Events to return</param>
        /// <returns></returns>
        private HttpResponseMessage eventResponse(List<Event> events)
        {
            var httpResponse = new HttpResponseMessage(HttpStatusCode.OK);

            MemoryStream stream = new MemoryStream();

            DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
            settings.UseSimpleDictionaryFormat = true;
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(List<Event>), settings);

            serializer.WriteObject(stream, events);

            httpResponse.Content = new StringContent(Encoding.UTF8.GetString(stream.ToArray()), Encoding.UTF8, "application/json");
            return httpResponse;
        }

        private HttpResponseMessage connectorResponse(string type)
        {
            return connectorResponse(new string[] { type });
        }

        /// <summary>
        /// Access a private instance variable
        /// </summary>
        /// <param name="type">Type of object that owns the private variable</param>
        /// <param name="instance">Object that owns the private variable</param>
        /// <param name="fieldName">Name of the private variable you wish to access</param>
        /// <returns>Private variable from your object</returns>
        private object getPrivateValue(Type type, object instance, string fieldName)
        {
            BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic
                | BindingFlags.Static;
            FieldInfo field = type.GetField(fieldName, bindFlags);
            return field.GetValue(instance);
        }
    }
}
