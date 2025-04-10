using System.Collections.Generic;

namespace ECATPlugin
{
    public class MasonryECData
    {
        public string Material { get; set; }
        public string ModuleA1A3 { get; set; }
        public string ModuleA4 { get; set; }
        public string ModuleA5 { get; set; }
        public string Density { get; set; }
    }

    public static class MasonryECDataProvider
    {
        public static List<MasonryECData> GetBlockworkData()
        {
            return new List<MasonryECData>
            {
                new MasonryECData {
                    Material = "Blockwork",
                    ModuleA1A3 = "0.093",
                    ModuleA4 = "0.032",
                    ModuleA5 = "0.250",
                    Density = "2000"
                }
            };
        }

        public static List<MasonryECData> GetBrickworkData()
        {
            return new List<MasonryECData>
            {
                new MasonryECData {
                    Material = "Brickwork",
                    ModuleA1A3 = "0.213",
                    ModuleA4 = "0.032",
                    ModuleA5 = "0.250",
                    Density = "1910"
                }
            };
        }

        public static List<MasonryECData> GetDataForMasonryType(string masonryType)
        {
            switch (masonryType)
            {
                case "Blockwork":
                    return GetBlockworkData();
                case "Brickwork":
                    return GetBrickworkData();
                default:
                    return new List<MasonryECData>();
            }
        }
    }
}