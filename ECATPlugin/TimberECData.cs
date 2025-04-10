using System.Collections.Generic;

public class TimberECData
{
    public string Source { get; set; }
    public string ModuleA1A3 { get; set; }
    public string ModuleA4 { get; set; }
    public string ModuleA5 { get; set; }
}

public static class TimberECDataProvider
{
    public static List<TimberECData> GetSoftwoodData()
    {
        return new List<TimberECData>
        {
            new TimberECData { Source = "Global", ModuleA1A3 = "0.263", ModuleA4 = "0.161", ModuleA5 = "0.0111" }
        };
    }

    public static List<TimberECData> GetGlulamData()
    {
        return new List<TimberECData>
        {
            new TimberECData { Source = "UK & EU", ModuleA1A3 = "0.280", ModuleA4 = "0.161", ModuleA5 = "0.010" },
            new TimberECData { Source = "Global", ModuleA1A3 = "0.512", ModuleA4 = "0.161", ModuleA5 = "0.0111" }
        };
    }

    public static List<TimberECData> GetLVLData()
    {
        return new List<TimberECData>
        {
            new TimberECData { Source = "Global", ModuleA1A3 = "0.390", ModuleA4 = "0.161", ModuleA5 = "0.010" }
        };
    }

    public static List<TimberECData> GetCLTData()
    {
        return new List<TimberECData>
        {
            new TimberECData { Source = "Global", ModuleA1A3 = "0.420", ModuleA4 = "0.161", ModuleA5 = "0.010" }
        };
    }

    public static List<TimberECData> GetDataForTimberType(string timberType)
    {
        switch (timberType)
        {
            case "Softwood":
                return GetSoftwoodData();
            case "Glulam":
                return GetGlulamData();
            case "LVL":
                return GetLVLData();
            case "CLT":
                return GetCLTData();
            default:
                return GetSoftwoodData(); // Default to Softwood
        }
    }

    public static int GetDensityForTimberType(string timberType)
    {
        switch (timberType)
        {
            case "Softwood":
                return 380;
            case "Glulam":
                return 470;
            case "LVL":
                return 510;
            case "CLT":
                return 492;
            default:
                return 380; // Default to Softwood
        }
    }
}