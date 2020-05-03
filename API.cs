using System;
using System.Collections.Generic;

using System.Xml.Serialization;

namespace DatadumpTool
{
    public class DataDump { }
    
    [XmlRoot("REGIONS")]
    public class RegionDataDump : DataDump
    {
        public RegionDataDump()
        {
            Regions = new List<Region>();
        }

        [XmlElement("REGION", typeof(Region))]
        public List<Region> Regions;
    }

    [XmlRoot("NATIONS")]
    public class NationDataDump : DataDump
    {
        public NationDataDump()
        {
            Nations = new List<Nation>();
        }

        [XmlElement("NATION", typeof(Nation))]
        public List<Nation> Nations;
    }
}