using System.Collections.Generic;

namespace ECATPlugin
{
    public class SteelECData
    {
        public string Source { get; set; }
        public string ModuleA1A3 { get; set; }
        public string ModuleA4 { get; set; }
        public string ModuleA5 { get; set; }
    }

    public static class SteelECDataProvider
    {
        public static List<SteelECData> GetOpenSectionData()
        {
            return new List<SteelECData>
            {
                new SteelECData { Source = "UK", ModuleA1A3 = "1.740", ModuleA4 = "0.032", ModuleA5 = "0.010" },
                new SteelECData { Source = "Global", ModuleA1A3 = "1.550", ModuleA4 = "0.183", ModuleA5 = "0.010" },
                new SteelECData { Source = "UK Reused", ModuleA1A3 = "0.050", ModuleA4 = "0.032", ModuleA5 = "0.010" }
            };
        }

        public static List<SteelECData> GetClosedSectionData()
        {
            return new List<SteelECData>
            {
                new SteelECData { Source = "UK", ModuleA1A3 = "2.50", ModuleA4 = "0.032", ModuleA5 = "0.010" },
                new SteelECData { Source = "Global", ModuleA1A3 = "2.50", ModuleA4 = "0.183", ModuleA5 = "0.010" }
            };
        }

        public static List<SteelECData> GetPlatesData()
        {
            return new List<SteelECData>
            {
                new SteelECData { Source = "UK", ModuleA1A3 = "2.46", ModuleA4 = "0.032", ModuleA5 = "0.010" },
                new SteelECData { Source = "Global", ModuleA1A3 = "2.46", ModuleA4 = "0.183", ModuleA5 = "0.010" }
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