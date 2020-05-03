using System;
using System.Xml.Serialization;

namespace DatadumpTool
{
    [Serializable()]
    public sealed class Officer
    {
        [XmlElement("NATION")]
        public readonly string Nation;
        [XmlElement("OFFICE")]
        public readonly string Office;
        [XmlElement("AUTHORITY")]
        public readonly string OfficerAuth;
        [XmlElement("TIME")]
        public readonly int AssingedTimestamp;
        [XmlElement("BY")]
        public readonly string AssignedBy;
        [XmlElement("ORDER")]
        public readonly int Order;
    }
}