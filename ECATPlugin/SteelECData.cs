using System.Collections.Generic;

namespace ECATPlugin
{
    public class SteelECData
    {
        public string Source { get; set; }
        public string ModuleA1A3 { get; set; }
        public string ModuleA1A5 { get; set; }
    }

    public static class SteelECDataProvider
    {
        public static List<SteelECData> GetOpenSectionData()
        {
            return new List<SteelECData>
            {
                new SteelECData { Source = "UK", ModuleA1A3 = "1.740", ModuleA1A5 = "1.782" },
                new SteelECData { Source = "Global", ModuleA1A3 = "1.550", ModuleA1A5 = "1.743" },
                new SteelECData { Source = "UK Reused", ModuleA1A3 = "0.050", ModuleA1A5 = "0.092" }
            };
        }

        public static List<SteelECData> GetClosedSectionData()
        {
            return new List<SteelECData>
            {
                new SteelECData { Source = "UK", ModuleA1A3 = "2.50", ModuleA1A5 = "2.542" },
                new SteelECData { Source = "Global", ModuleA1A3 = "2.50", ModuleA1A5 = "2.693" }
            };
        }

        public static List<SteelECData> GetPlatesData()
        {
            return new List<SteelECData>
            {
                new SteelECData { Source = "UK", ModuleA1A3 = "2.46", ModuleA1A5 = "2.502" },
                new SteelECData { Source = "Global", ModuleA1A3 = "2.46", ModuleA1A5 = "2.653" }
            };
        }

        public static List<SteelECData> GetDataForSectionType(string sectionType)
        {
            switch (sectionType)
            {
                case "Open Sections":
                    return GetOpenSectionData();
                case "Closed Sections":
                    return GetClosedSectionData();
                case "Plates":
                    return GetPlatesData();
                default:
                    return GetOpenSectionData(); // Default to Open Sections
            }
        }
    }
}