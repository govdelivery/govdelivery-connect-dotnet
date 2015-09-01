using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace HubAgent
{
    class Program
    {
        static void Main(string[] args)
        {
            // Grab configuration
            var dynamicsHost = ConfigurationManager.AppSettings["DynamicsHost"].ToString();
            var dynamicsUsername = ConfigurationManager.AppSettings["DynamicsUsername"].ToString();
            var dynamicsPassword = ConfigurationManager.AppSettings["DynamicsPassword"].ToString();

            var govdeliveryHost = ConfigurationManager.AppSettings["GovDeliveryHost"].ToString();
            var govdeliveryKey = ConfigurationManager.AppSettings["GovDeliveryKey"].ToString();

            var dynamics = new DynamicsClient(dynamicsHost, dynamicsUsername, dynamicsPassword);
            var hub = new HubClient(govdeliveryHost, govdeliveryKey, dynamics);

            // Initialize connection information for hub
            hub.RetrieveConnectors().Wait();

            // Ensure that dynamics is properly setup for our use case
            dynamics.EnsureGovdeliveryPublisher();
            dynamics.EnsureGovdeliveryEmailFields();
            dynamics.EnsureEmailMetadata();

            // Retrieve and send emails that are in "Pending Send"
            var emails = dynamics.RetrievePendingEmails(hub.GetEmailForPolling());
            hub.SendEmails(emails).Wait();

            // Push analytics data back into dynamics
            hub.UpdateEmailStatuses().Wait();
        }
    }
}