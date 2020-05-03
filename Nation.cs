using System;
using System.Xml.Serialization;

namespace DatadumpTool
{
    /// <summary>
    /// A container for nation data-dump items
    /// We only save the things relevant to raiding
    /// </summary>
    /// 
    [Serializable()]
    public sealed class Nation
    {
        [XmlElement("NAME")]
        public string name;
        public string Name
        {
            get
            {
                return name.Replace(' ','_').ToLower();
            }
        }

        [XmlElement("UNSTATUS")]
        public string WAStatus;
        [XmlElement("ENDORSEMENTS")]
        public string Endorsements;
        [XmlElement("REGION")]
        public string Region;

        //Elements added by ARCore
        public int Index;
    }
}