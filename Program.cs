using System;
using System.IO;
using System.Linq;

using System.Text;
using System.Xml.Serialization;

using System.IO.Compression;
using System.Threading.Tasks;
using System.Collections.Generic;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.ConditionalFormatting;
using System.Drawing;

using ShellProgressBar;

namespace DatadumpTool
{
    class Program
    {
        private string UnzipDump(string Filename)
        {
            //This will contain the raw XML
            string text;
            //These usings initialize the filestream of the downloaded naiton data dump, and the
            //GZip stream that will decompress it. They'll both be disposed afterwads
            using (var FS = new FileStream(Filename, FileMode.Open))
            using (var GZip = new GZipStream(FS, CompressionMode.Decompress))
            {
                //Initialize the buffer
                const int size = 4096;
                byte[] buffer = new byte[size];
                //Initialize the memory stream that will contain the decompressed object
                using (MemoryStream memory = new MemoryStream())
                {
                    //Read bytes into the buffer until there is nothing left to read.
                    int count = 0;
                    do
                    {
                        count = GZip.Read(buffer, 0, size);
                        if (count > 0)
                        {
                            memory.Write(buffer, 0, count);
                        }
                    }
                    while (count > 0);
                    //convert the decompressed bytes into a string.
                    byte[] decompressed = memory.ToArray();
                    text = Encoding.ASCII.GetString(decompressed);
                }
            }

            return text;
        }

        public T ParseDataDump<T>(string Filename) where T : DataDump
        {
            string RawXML;

            //If it's XML, don't bother unzipping it
            if (Filename.EndsWith(".xml"))
                RawXML = File.ReadAllText(Filename); //
            else if (Filename.EndsWith(".xml.gz"))
                RawXML = UnzipDump(Filename);
            else
                throw new ArgumentException("Invalid data-dump file. Must be an XML file, or GZipped XML file.");

            //Deserialize the XML into the DataDump object
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            using var fs = new StringReader(RawXML);
            T dump = (T)serializer.Deserialize(fs);

            return dump;
        }

        public async Task MainAsync(string[] args)
        {
            using var MasterBar = new ProgressBar(5, "Overall Progress");

            using var LoadBar = MasterBar.Spawn(2, "Loading data dumps.");
            var RegionsBefore = ParseDataDump<RegionDataDump>($"./dumps/{args[0]}-regions.xml.gz").Regions;
            LoadBar.Tick();
            var RegionsAfter = ParseDataDump<RegionDataDump>($"./dumps/{args[1]}-regions.xml.gz").Regions;
            LoadBar.Tick();
            List<Nation> Nations1 = ParseDataDump<NationDataDump>($"./dumps/{args[0]}-nations.xml.gz").Nations;
            LoadBar.Tick();
            List<Nation> Nations2 = ParseDataDump<NationDataDump>($"./dumps/{args[1]}-nations.xml.gz").Nations;
            LoadBar.Tick();
            MasterBar.Tick();

            //Create a region list
            using var RegionsList = MasterBar.Spawn(Nations1.Count+Nations2.Count, "Generating region list");
            List<string> Regions = new List<string>();
            Nations1.ForEach(N=>{
                if(!Regions.Contains(N.Region))
                    Regions.Add(N.Region);
                RegionsList.Tick();
            });
            Nations2.ForEach(N=>{
                if(!Regions.Contains(N.Region))
                    Regions.Add(N.Region);
                RegionsList.Tick();
            });
            MasterBar.Tick();

            //Populate the endorsement and WA Member maps
            using var EndoCounting1 = MasterBar.Spawn(Nations1.Count, "Counting up endorsements from first dump.");
            Dictionary<string, int> EndoMap1 = new Dictionary<string, int>();
            Dictionary<string, int> WAMembermap1 = new Dictionary<string, int>();
            foreach(Nation nation in Nations1)           
            {
                if(nation.Endorsements == "")
                {
                    EndoCounting1.Tick();
                    continue;
                }
                
                //endoshit
                int endoCount = nation.Endorsements.Split(",").Length;
                if(!EndoMap1.ContainsKey(nation.Region))
                    EndoMap1.Add(nation.Region, 0);
                EndoMap1[nation.Region] += endoCount;
                //WAShit
                if(!WAMembermap1.ContainsKey(nation.Region))
                    WAMembermap1.Add(nation.Region, 0);
                if(nation.WAStatus == "WA Member")
                    WAMembermap1[nation.Region]++;
                EndoCounting1.Tick();
            }
            MasterBar.Tick();

            using var PopProgress = MasterBar.Spawn(RegionsBefore.Count+RegionsBefore.Count, "Couting up regional populations.");
            Dictionary<string, int> BeforeMap = new Dictionary<string, int>();
            RegionsBefore.ForEach(R=>{
                BeforeMap.Add(R.name, R.NumNations);
                PopProgress.Tick();
            });

            Dictionary<string, int> AfterMap = new Dictionary<string, int>();
            RegionsAfter.ForEach(R=>{
                AfterMap.Add(R.name, R.NumNations);
                PopProgress.Tick();
            });
            MasterBar.Tick();

            using var EndoCounting2 = MasterBar.Spawn(Nations2.Count, "Counting up endorsements from second dump.");
            Dictionary<string, int> EndoMap2 = new Dictionary<string, int>();
            Dictionary<string, int> WAMembermap2 = new Dictionary<string, int>();
            foreach(Nation nation in Nations2)           
            {
                if(nation.Endorsements == "")
                {
                    EndoCounting2.Tick();
                    continue;
                }
                
                int endoCount = nation.Endorsements.Split(",").Length;
                if(!EndoMap2.ContainsKey(nation.Region))
                    EndoMap2.Add(nation.Region, 0);
                EndoMap2[nation.Region] += endoCount;
                //WAShit
                if(!WAMembermap2.ContainsKey(nation.Region))
                    WAMembermap2.Add(nation.Region, 0);
                if(nation.WAStatus == "WA Member")
                    WAMembermap2[nation.Region]++;
                EndoCounting2.Tick();
            }
            MasterBar.Tick();

            //Initialize the spreadsheet
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Post-Drew Regional Growth");
            worksheet.Cells[1,1].Value = "Region"; //A

            worksheet.Cells[1,2].Value = "Initial Pop."; //B
            worksheet.Cells[1,3].Value = "New Pop."; //C

            worksheet.Cells[1,4].Value = "Initial WA Count"; //D
            worksheet.Cells[1,5].Value = "New WA Count"; //E

            worksheet.Cells[1,6].Value = "Pop Delta"; //F
            worksheet.Cells[1,7].Value = "% Change"; //G

            worksheet.Cells[1,8].Value = "WA Delta"; //H
            worksheet.Cells[1,9].Value = "% Change"; //I

            //conditional rules for pop delta
            var Address = new ExcelAddress("F2:F30000");
            var GreaterRule = worksheet.ConditionalFormatting.AddGreaterThan(Address);
            GreaterRule.Formula = "0";
            GreaterRule.Style.Fill.PatternType = ExcelFillStyle.Solid;
            GreaterRule.Style.Fill.BackgroundColor.Color = Color.LimeGreen;

            var LesserRule = worksheet.ConditionalFormatting.AddLessThan(Address);
            LesserRule.Formula = "0";
            LesserRule.Style.Fill.PatternType = ExcelFillStyle.Solid;
            LesserRule.Style.Fill.BackgroundColor.Color = Color.Salmon;

            //Rules for the WA Delta
            var AddressWADelta = new ExcelAddress("H2:H30000");
            var GreaterRuleWADelta = worksheet.ConditionalFormatting.AddGreaterThan(AddressWADelta);
            GreaterRuleWADelta.Formula = "0";
            GreaterRuleWADelta.Style.Fill.PatternType = ExcelFillStyle.Solid;
            GreaterRuleWADelta.Style.Fill.BackgroundColor.Color = Color.LimeGreen;

            var LesserRuleWADelta = worksheet.ConditionalFormatting.AddLessThan(AddressWADelta);
            LesserRuleWADelta.Formula = "0";
            LesserRuleWADelta.Style.Fill.PatternType = ExcelFillStyle.Solid;
            LesserRuleWADelta.Style.Fill.BackgroundColor.Color = Color.Salmon;
 
            //Do the work
            using var MainProgress = MasterBar.Spawn(Regions.Count*2, "Populating spreadsheet.");
            for(int i = 0; i < Regions.Count; i++)
            {
                string Region = Regions[i];
                worksheet.Cells[i+2,1].Value = Region;

                //Population raw data
                if(BeforeMap.ContainsKey(Region))
                    worksheet.Cells[i+2,2].Value = BeforeMap[Region];
                else
                    worksheet.Cells[i+2,2].Value = 0;

                if(AfterMap.ContainsKey(Region))
                    worksheet.Cells[i+2,3].Value = AfterMap[Region];
                else
                    worksheet.Cells[i+2,3].Value = 0;

                //WA raw data
                if(WAMembermap1.ContainsKey(Region))
                    worksheet.Cells[i+2,4].Value = WAMembermap1[Region];
                else
                    worksheet.Cells[i+2,4].Value = 0;

                if(WAMembermap2.ContainsKey(Region))
                    worksheet.Cells[i+2,5].Value = WAMembermap2[Region];
                else
                    worksheet.Cells[i+2,5].Value = 0;
                MainProgress.Tick();

                //Population Delta fomula
                worksheet.Cells[i+2,6].Formula = $"C{i+2}-B{i+2}";
                worksheet.Cells[i+2,7].Formula = $"(C{i+2}-B{i+2})/B{i+2}";
                worksheet.Cells[i+2,7].Style.Numberformat.Format = "#0.00%";

                //WA Delta formula
                worksheet.Cells[i+2,8].Formula = $"E{i+2}-D{i+2}";
                worksheet.Cells[i+2,9].Formula = $"(E{i+2}-D{i+2})/D{i+2}";
                worksheet.Cells[i+2,9].Style.Numberformat.Format = "#0.00%";
                MainProgress.Tick();
            }
            MasterBar.Tick();
            
            worksheet.Cells.AutoFitColumns(0);
            worksheet.Calculate();
            //Save the spreadsheet
            if(!Directory.Exists("./sheets"))
                Directory.CreateDirectory("./sheets");

            string filename = $"./sheets/{args[0]}-{args[1]} PopCompare.xlsx";
            try{File.Delete(filename);}catch(Exception){}
            using (var fs = new FileStream(filename, FileMode.CreateNew, FileAccess.Write, FileShare.None)){
                package.SaveAs(fs);
            }
        }

        static void Main(string[] args)
        {
            new Program().MainAsync(args).GetAwaiter().GetResult();
        }
    }
}
