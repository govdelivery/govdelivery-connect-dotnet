using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace HubAgent
{
    class HubClient
    {
        private HttpClient connection;
        private IDynamicsClient dynamicsClient;
        private List<Connector> connectors;
        private Connector dynamics;

        private readonly List<string> transitions = new List<string>() { "sent", "opened", "failed", "bounced" };
        private readonly List<string> statuses = new List<string>() { "Sent", "Received", "Failed", "Failed" };

        public HubClient(string url, string key, IDynamicsClient client)
        {
            dynamicsClient = client;
            connection = new HttpClient(new LoggingHandler(new HttpClientHandler()));
            connection.MaxResponseContentBufferSize = 256000;
            connection.DefaultRequestHeaders.Add("user-agent", "Hub.NET-Agent");
            connection.DefaultRequestHeaders.Add("accept", "application/json");
            connection.DefaultRequestHeaders.Add("x-agent-token", key);
            connection.BaseAddress = new Uri(url);
        }

        /// <summary>
        /// Get the email address used in polling dynamics for emails to send through govdelivery
        /// </summary>
        /// <returns></returns>
        public string GetEmailForPolling()
        {
            return dynamics.options["govdelivery_dynamics_email"];
        }

        /// <summary>
        /// Retrieve connectors for the client
        /// This method MUST be called on any HubClient instance prior to use
        /// </summary>
        /// <returns></returns>

        public async Task RetrieveConnectors()
        {
            HttpResponseMessage response = await connection.GetAsync("connectors");
            if (response.IsSuccessStatusCode)
            {
                MemoryStream data = (MemoryStream)(await response.Content.ReadAsStreamAsync());

                DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                settings.UseSimpleDictionaryFormat = true;

                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(List<Connector>), settings);
                connectors = (List<Connector>)serializer.ReadObject(data);
                foreach (Connector conn in connectors)
                {
                    if (conn.type == "Connectors::DynamicsEmail")
                    {
                        dynamics = conn;
                    }
                }
            }
        }

        /// <summary>
        /// Grab a list of emails to send from dynamics and push up to hub
        /// </summary>
        /// <returns></returns>
        public async Task SendEmails(List<Email> emails)
        {
            foreach (Link link in dynamics.links)
            {
                if (link.rel == "actions")
                {
                    String endpoint = link.href + "/deliver";

                    foreach (Email email in emails)
                    {
                        MemoryStream stream = new MemoryStream();
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(EmailPayload));
                        serializer.WriteObject(stream, new EmailPayload() { payload = email });

                        StringContent content = new StringContent(Encoding.UTF8.GetString(stream.ToArray()), Encoding.UTF8, "application/json");
                        HttpResponseMessage response = await connection.PostAsync(endpoint, content);
                        if (response.IsSuccessStatusCode)
                        {
                            dynamicsClient.UpdateStatus(email.id, "Sending");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Grab a list of email status update events from hub and push back to dynamics
        /// </summary>
        /// <returns></returns>
        public async Task UpdateEmailStatuses()
        {
            foreach (Link link in dynamics.links)
            {
                if (link.rel == "actions")
                {
                    String endpoint = link.href + "/statuses";
                    HttpResponseMessage response = await connection.GetAsync(endpoint);
                    if (response.IsSuccessStatusCode)
                    {
                        MemoryStream data = (MemoryStream)(await response.Content.ReadAsStreamAsync());
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(List<Event>));
                        var events = (List<Event>)serializer.ReadObject(data);

                        Dictionary<int, int> messageStatuses = new Dictionary<int, int>();
                        foreach (Event ev in events)
                        {
                            if (ev.name == "email_sent")
                            {
                                dynamicsClient.AssociateGovdeliveryEmail(ev.external_id, ev.message_id.ToString());
                            }
                            else if (ev.name == "email_status")
                            {
                                var statusIndex = transitions.IndexOf(ev.status);
                                if ((statusIndex > -1) && (!messageStatuses.ContainsKey(ev.message_id) ||
                                    ( statusIndex > messageStatuses[ev.message_id]))) {
                                        messageStatuses[ev.message_id] = statusIndex;
                                }
                            }
                        }
                        foreach (KeyValuePair<int, int> message in messageStatuses)
                        {
                            var dynamicsId = dynamicsClient.LookupEmailByGovdeliveryId(message.Key.ToString());
                            if (dynamicsId != null)
                            {
                                dynamicsClient.UpdateStatus(dynamicsId, statuses[message.Value]);
                            }
                        }
                    }
                }
            }
        }
    }
}
