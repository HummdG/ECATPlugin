using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ECATPlugin.Helpers
{
    public static class RevitHelper
    {
        public static ElementId GetMaterialIdFromName(Document doc, string materialName)
        {
            var materials = new FilteredElementCollector(doc)
                .OfClass(typeof(Material))
                .ToElements();

            foreach (Material mat in materials)
            {
                if (mat.Name == materialName)
                {
                    return mat.Id;
                }
            }

            return null;
        }

        public static List<Element> GetAllElementsOfCategory(Document doc, BuiltInCategory category)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();
        }

        public static List<Element> GetAllStructuralElements(Document doc, StructuralType structuralType)
        {
            var filter = new ElementStructuralTypeFilter(structuralType);
            return new FilteredElementCollector(doc)
                .WherePasses(filter)
                .ToElements()
                .ToList();
        }

        // In RevitHelper.cs, add this new method
        public static ElementId GetMaterialIdByFlexibleName(Document doc, string materialName)
        {
            var materials = new FilteredElementCollector(doc)
                .OfClass(typeof(Material))
                .ToElements()
                .Cast<Material>();

            // First try exact match
            var exactMatch = materials.FirstOrDefault(m => m.Name.Equals(materialName, StringComparison.OrdinalIgnoreCase));
            if (exactMatch != null)
                return exactMatch.Id;

            // Then try contains match
            var containsMatch = materials.FirstOrDefault(m => m.Name.IndexOf(materialName, StringComparison.OrdinalIgnoreCase) >= 0);
            if (containsMatch != null)
                return containsMatch.Id;

            // Try checking for common alternatives
            if (materialName == "WAL_Steel")
            {
                var steelMatch = materials.FirstOrDefault(m => m.Name.IndexOf("Steel", StringComparison.OrdinalIgnoreCase) >= 0);
                if (steelMatch != null)
                    return steelMatch.Id;
            }
            else if (materialName == "Wood - Dimensional Lumber")
            {
                var woodMatch = materials.FirstOrDefault(m =>
                    m.Name.IndexOf("Wood", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Name.IndexOf("Timber", StringComparison.OrdinalIgnoreCase) >= 0);
                if (woodMatch != null)
                    return woodMatch.Id;
            }
            else if (materialName == "WAL_Block" || materialName == "WAL_Brick")
            {
                var masonryMatch = materials.FirstOrDefault(m =>
                    m.Name.IndexOf("Block", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Name.IndexOf("Brick", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Name.IndexOf("Masonry", StringComparison.OrdinalIgnoreCase) >= 0);
                if (masonryMatch != null)
                    return masonryMatch.Id;
            }

            return null;
        }
    }
}