using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace HubAgent
{
    [DataContract]
    internal class Event
    {
        [DataMember(Name="event")]
        internal string name;
        [DataMember]
        internal string external_id;
        [DataMember]
        internal int message_id;
        [DataMember]
        internal string status;
        [DataMember]
        internal string error_message;
    }

    [DataContract]
    internal class Link
    {
        [DataMember]
        internal string rel;
        [DataMember]
        internal string href;
    }

    [DataContract]
    internal class Connector
    {
        [DataMember]
        internal int id;
        [DataMember]
        internal string type;
        [DataMember]
        internal Dictionary<string, string> options;
        [DataMember]
        internal List<Link> links;
    }

    [DataContract]
    public class Email
    {
        [DataMember(Name = "external_id")]
        internal string id;
        [DataMember(Name = "from_email")]
        internal string from;
        [DataMember]
        internal List<string> to;
        [DataMember]
        internal string subject;
        [DataMember]
        internal string body;
    }

    [DataContract]
    internal class EmailPayload
    {
        [DataMember]
        internal Email payload;
    }
}
