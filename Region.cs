using System;
using System.Collections.Generic;

using System.Xml;
using System.Xml.Serialization;

namespace DatadumpTool
{
        [Serializable()]
    public sealed class Region
    {
        //These are values parsed from the data dump
        [XmlElement("NAME")]
        public string name;
        public string Name
        {
            get
            {
                return name.Replace(' ','_').ToLower();
            }
        }

        [XmlElement("NUMNATIONS")]
        public int NumNations;
        [XmlElement("NATIONS")]
        public string nations;
        [XmlElement("DELEGATE")]
        public string Delegate;
        [XmlElement("DELEGATEVOTES")]
        public int DelegateVotes;
        [XmlElement("DELEGATEAUTH")]
        public string DelegateAuth;
        [XmlElement("FOUNDER")]
        public string Founder;
        [XmlElement("FOUNDERAUTH")]
        public string FounderAuth;
        [XmlArray("OFFICERS"), XmlArrayItem("OFFICER", typeof(Officer))]
        public List<Officer> Officers;
        [XmlArray("EMBASSIES"), XmlArrayItem("EMBASSY", typeof(string))]
        public string[] Embassies;
        [XmlElement("LASTUPDATE")]
        public double lastUpdate;
        public double LastUpdate
        {
                get
                {
                    //Subtract 4 hours from LastUpdate
                    //Seconds into the update is more useful than the UTC update time
                    return lastUpdate - (4 *3600);
                }
        }

        //These are values added after the fact by ARCore
        public string[] Nations
        {
            get {
                return nations
                    .Replace('_',' ')
                    .Split(":", StringSplitOptions.RemoveEmptyEntries);
            }
        }

        public int Index;
        public bool hasPassword;
        public bool hasFounder;
    }
}