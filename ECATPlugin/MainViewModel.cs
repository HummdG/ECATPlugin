using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using ECATPlugin.Helpers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Media;
using System.Windows.Input;
using System.Xml.Linq;

namespace ECATPlugin
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private Document _document;
        private string _projectName;
        private string _projectNumber;
        public ICommand ShowRatingsCommand { get; private set; }

        public MainViewModel(Document doc)
        {
            _document = doc;
            InitializeItems();
            InitializeMaterialData();
            LoadVolumes();
            LoadGIA();
            GetProjectName();
            GetProjectNumber();
            UpdateCarbonRatingColor();
            LoadData(); // Load saved data
            ShowRatingsCommand = new RelayCommand(ShowOtherRatings);
        }

        private void ShowOtherRatings()
        {
            // Create an instance of RatingPopup with the CarbonRating value
            RatingPopup ratingPopup = new RatingPopup(CarbonRating, this);
            ratingPopup.ShowDialog(); // Show the popup as a dialog
        }

        // Define structure to hold item properties
        public class ItemData : INotifyPropertyChanged
        {
            private double _volume;
            public double Volume
            {
                get => _volume;
                set
                {
                    _volume = value;
                    OnPropertyChanged(nameof(Volume));
                    OnPropertyChanged(nameof(SubTotalCarbon));
                }
            }

            private double _concreteVolume;
            private double _steelVolume;
            private double _timberVolume;
            private double _masonryVolume;

            public double ConcreteVolume
            {
                get => _concreteVolume;
                set
                {
                    _concreteVolume = value;
                    OnPropertyChanged(nameof(ConcreteVolume));
                    UpdateTotalVolume();
                }
            }

            public double SteelVolume
            {
                get => _steelVolume;
                set
                {
                    _steelVolume = value;
                    OnPropertyChanged(nameof(SteelVolume));
                    UpdateTotalVolume();
                }
            }

            public double TimberVolume
            {
                get => _timberVolume;
                set
                {
                    _timberVolume = value;
                    OnPropertyChanged(nameof(TimberVolume));
                    UpdateTotalVolume();
                }
            }

            public double MasonryVolume
            {
                get => _masonryVolume;
                set
                {
                    _masonryVolume = value;
                    OnPropertyChanged(nameof(MasonryVolume));
                    UpdateTotalVolume();
                }
            }

            private void UpdateTotalVolume()
            {
                Volume = ConcreteVolume + SteelVolume + TimberVolume + MasonryVolume;
            }

            private double _ec;
            public double EC
            {
                get => _ec;
                set
                {
                    _ec = value;
                    OnPropertyChanged(nameof(EC));
                    OnPropertyChanged(nameof(SubTotalCarbon));
                }
            }

            private double _rebarDensity;
            public double RebarDensity
            {
                get => _rebarDensity;
                set
                {
                    _rebarDensity = value;
                    OnPropertyChanged(nameof(RebarDensity));
                    OnPropertyChanged(nameof(SubTotalCarbon));
                }
            }

            // Steel specific properties
            public string SteelType { get; set; } = "Open Section"; // Open Section, Closed Section, Plates
            public string SteelSource { get; set; } = "UK"; // UK, Global, UK (Reused)

            // Timber specific properties
            public string TimberType { get; set; } = "Softwood"; // Softwood, Glulam, LVL, CLT
            public string TimberSource { get; set; } = "Global"; // UK & EU, Global

            // Masonry specific properties
            public string MasonryType { get; set; } = "Blockwork"; // Blockwork, Brickwork

            // Module properties for all materials
            public string Module { get; set; } = "A1-A3"; // A1-A3, A4, A5

            public double SubTotalCarbon => CalculateSubTotalCarbon();

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

            private double CalculateSubTotalCarbon()
            {
                // Base calculation for concrete
                double concreteCarbon = (ConcreteVolume * EC + ConcreteVolume * RebarDensity) * 0.785;

                // Calculate for steel based on type and source
                double steelCarbon = CalculateSteelCarbon();

                // Calculate for timber based on type and source
                double timberCarbon = CalculateTimberCarbon();

                // Calculate for masonry based on type
                double masonryCarbon = CalculateMasonryCarbon();

                return concreteCarbon + steelCarbon + timberCarbon + masonryCarbon;
            }

            private double CalculateSteelCarbon()
            {
                if (SteelVolume <= 0) return 0;

                double ecFactor = GetSteelECFactor();
                return SteelVolume * ecFactor * 7850; // Density of steel = 7850 kg/m³
            }

            private double GetSteelECFactor()
            {
                // Values from the provided PDF
                if (SteelType == "Open Section")
                {
                    if (SteelSource == "UK") return 1.740;
                    if (SteelSource == "Global") return 1.550;
                    if (SteelSource == "UK (Reused)") return 0.050;
                }
                else if (SteelType == "Closed Section")
                {
                    if (SteelSource == "UK") return 2.500;
                    if (SteelSource == "Global") return 2.500;
                }
                else if (SteelType == "Plates")
                {
                    if (SteelSource == "UK") return 2.460;
                    if (SteelSource == "Global") return 2.460;
                }

                return 1.740; // Default to UK Open Section
            }

            private double CalculateTimberCarbon()
            {
                if (TimberVolume <= 0) return 0;

                double ecFactor = GetTimberECFactor();
                double density = GetTimberDensity();

                return TimberVolume * ecFactor * density;
            }

            private double GetTimberECFactor()
            {
                // Values from the provided PDF
                if (TimberType == "Softwood")
                {
                    return 0.263;
                }
                else if (TimberType == "Glulam")
                {
                    if (TimberSource == "UK & EU") return 0.280;
                    if (TimberSource == "Global") return 0.512;
                }
                else if (TimberType == "LVL")
                {
                    return 0.390;
                }
                else if (TimberType == "CLT")
                {
                    return 0.420;
                }

                return 0.263; // Default to Softwood Global
            }

            private double GetTimberDensity()
            {
                // Densities from the provided PDF
                if (TimberType == "Softwood") return 380;
                if (TimberType == "Glulam") return 470;
                if (TimberType == "LVL") return 510;
                if (TimberType == "CLT") return 492;

                return 380; // Default to Softwood
            }

            private double CalculateMasonryCarbon()
            {
                if (MasonryVolume <= 0) return 0;

                double ecFactor = 0;
                double density = 0;

                if (MasonryType == "Blockwork")
                {
                    ecFactor = 0.093;
                    density = 2000;
                }
                else if (MasonryType == "Brickwork")
                {
                    ecFactor = 0.213;
                    density = 1910;
                }

                return MasonryVolume * ecFactor * density;
            }
        }

        // Dictionary to store item data
        private Dictionary<string, ItemData> _items = new Dictionary<string, ItemData>();

        // Material data dictionaries
        private Dictionary<string, Dictionary<string, double>> _steelData = new Dictionary<string, Dictionary<string, double>>();
        private Dictionary<string, Dictionary<string, double>> _timberData = new Dictionary<string, Dictionary<string, double>>();
        private Dictionary<string, Dictionary<string, double>> _masonryData = new Dictionary<string, Dictionary<string, double>>();

        // ItemsView for binding to DataGrid
        private ObservableCollection<KeyValuePair<string, ItemData>> _itemsView;
        public ObservableCollection<KeyValuePair<string, ItemData>> ItemsView
        {
            get
            {
                if (_itemsView == null)
                {
                    _itemsView = new ObservableCollection<KeyValuePair<string, ItemData>>();
                    foreach (var item in _items)
                    {
                        _itemsView.Add(item);
                    }
                }
                return _itemsView;
            }
        }

        // Expose the dictionary as a public property
        public Dictionary<string, ItemData> Items
        {
            get => _items;
            set
            {
                _items = value;
                OnPropertyChanged(nameof(Items));
                UpdateItemsView();
            }
        }

        // Method to update ItemsView when the Items dictionary changes
        private void UpdateItemsView()
        {
            _itemsView = null;
            OnPropertyChanged(nameof(ItemsView));
        }

        // Initialize dictionary with default values
        private void InitializeItems()
        {
            _items["Beam"] = new ItemData { EC = 250, RebarDensity = 225 };
            _items["Wall"] = new ItemData { EC = 250, RebarDensity = 180 };
            _items["Upstand"] = new ItemData { EC = 250, RebarDensity = 180 };
            _items["Column"] = new ItemData { EC = 250, RebarDensity = 250 };
            _items["Floor"] = new ItemData { EC = 215, RebarDensity = 120 };
            _items["Foundation"] = new ItemData { EC = 215, RebarDensity = 110 };
            _items["Piling"] = new ItemData { EC = 140, RebarDensity = 75 };

            foreach (var item in _items.Values)
            {
                item.PropertyChanged += ItemData_PropertyChanged;
            }
        }

        private void InitializeMaterialData()
        {
            // Initialize steel data
            _steelData["Open Section"] = new Dictionary<string, double>
            {
                { "UK", 1.740 },
                { "Global", 1.550 },
                { "UK (Reused)", 0.050 }
            };

            _steelData["Closed Section"] = new Dictionary<string, double>
            {
                { "UK", 2.500 },
                { "Global", 2.500 }
            };

            _steelData["Plates"] = new Dictionary<string, double>
            {
                { "UK", 2.460 },
                { "Global", 2.460 }
            };

            // Initialize timber data
            _timberData["Softwood"] = new Dictionary<string, double>
            {
                { "Global", 0.263 }
            };

            _timberData["Glulam"] = new Dictionary<string, double>
            {
                { "UK & EU", 0.280 },
                { "Global", 0.512 }
            };

            _timberData["LVL"] = new Dictionary<string, double>
            {
                { "Global", 0.390 }
            };

            _timberData["CLT"] = new Dictionary<string, double>
            {
                { "Global", 0.420 }
            };

            // Initialize masonry data
            _masonryData["Blockwork"] = new Dictionary<string, double>
            {
                { "A1-A3", 0.093 },
                { "A4", 0.032 },
                { "A5", 0.250 }
            };

            _masonryData["Brickwork"] = new Dictionary<string, double>
            {
                { "A1-A3", 0.213 },
                { "A4", 0.032 },
                { "A5", 0.250 }
            };
        }

        private void ItemData_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // Update TotalCarbonEmbodiment if relevant properties change
            if (e.PropertyName == nameof(ItemData.Volume) ||
                e.PropertyName == nameof(ItemData.EC) ||
                e.PropertyName == nameof(ItemData.RebarDensity) ||
                e.PropertyName == nameof(ItemData.SubTotalCarbon))
            {
                UpdateTotalCarbonEmbodiment();
            }
        }

        public void Dispose()
        {
            SaveData(); // Save the data when disposing
            foreach (var item in _items.Values)
            {
                item.PropertyChanged -= ItemData_PropertyChanged;
            }
        }

        // Event to notify changes for data binding
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // Project name property
        public string ProjectName
        {
            get => _projectName;
            set
            {
                _projectName = value;
                OnPropertyChanged(nameof(ProjectName));
            }
        }

        // Project number property
        public string ProjectNumber
        {
            get => _projectNumber;
            set
            {
                _projectNumber = value;
                OnPropertyChanged(nameof(ProjectNumber));
            }
        }

        // Total embodied carbon property
        private double _totalCarbonEmbodiment;
        public double TotalCarbonEmbodiment
        {
            get => _totalCarbonEmbodiment;
            set
            {
                _totalCarbonEmbodiment = value;
                OnPropertyChanged(nameof(TotalCarbonEmbodiment));
            }
        }

        // Carbon rating property
        private double _carbonRating;
        public double CarbonRating
        {
            get => _carbonRating;
            set
            {
                _carbonRating = value;
                OnPropertyChanged(nameof(CarbonRating));
                UpdateCarbonRatingColor();
            }
        }

        // CarbonRatingColor property to bind to the background color of the Border
        private Brush _carbonRatingColor;
        public Brush CarbonRatingColor
        {
            get => _carbonRatingColor;
            private set
            {
                if (_carbonRatingColor != value)
                {
                    _carbonRatingColor = value;
                    OnPropertyChanged(nameof(CarbonRatingColor));
                }
            }
        }

        private string _carbonGrade;
        public string CarbonGrade
        {
            get => _carbonGrade;
            set
            {
                _carbonGrade = value;
                OnPropertyChanged(nameof(CarbonGrade));
            }
        }

        private void UpdateCarbonRatingColor()
        {
            if (CarbonRating <= 55)
            {
                CarbonRatingColor = Brushes.Green;         // A++
                CarbonGrade = "A++";
            }
            else if (CarbonRating <= 160)
            {
                CarbonRatingColor = Brushes.LimeGreen;    // A+
                CarbonGrade = "A+";
            }
            else if (CarbonRating <= 185)
            {
                CarbonRatingColor = Brushes.YellowGreen;  // A
                CarbonGrade = "A";
            }
            else if (CarbonRating <= 230)
            {
                CarbonRatingColor = Brushes.Yellow;       // B
                CarbonGrade = "B";
            }
            else if (CarbonRating <= 265)
            {
                CarbonRatingColor = Brushes.Gold;         // C
                CarbonGrade = "C";
            }
            else if (CarbonRating <= 320)
            {
                CarbonRatingColor = Brushes.Orange;       // D
                CarbonGrade = "D";
            }
            else if (CarbonRating <= 370)
            {
                CarbonRatingColor = Brushes.OrangeRed;    // E
                CarbonGrade = "E";
            }
            else if (CarbonRating <= 425)
            {
                CarbonRatingColor = Brushes.Red;          // F
                CarbonGrade = "F";
            }
            else
            {
                CarbonRatingColor = Brushes.DarkRed;      // G
                CarbonGrade = "G";
            }
        }

        // GIA properties
        private double _gia;
        public double GIA
        {
            get => _gia;
            set
            {
                _gia = value;
                OnPropertyChanged(nameof(GIA));
                UpdateCarbonRating();
            }
        }

        private double _giaCalculated;
        public double GIACalculated
        {
            get => _giaCalculated;
            set
            {
                _giaCalculated = value;
                OnPropertyChanged(nameof(GIACalculated));
                UpdateCarbonRating();
            }
        }

        // Properties to track which GIA is selected
        private bool _isInputGIASelected;
        public bool IsInputGIASelected
        {
            get => _isInputGIASelected;
            set
            {
                if (_isInputGIASelected != value)
                {
                    _isInputGIASelected = value;
                    OnPropertyChanged(nameof(IsInputGIASelected));
                    // If we're selecting input GIA, ensure calculated GIA is deselected
                    if (value && _isCalculatedGIASelected)
                    {
                        _isCalculatedGIASelected = false;
                        OnPropertyChanged(nameof(IsCalculatedGIASelected));
                    }
                    UpdateCarbonRating();
                }
            }
        }

        private bool _isCalculatedGIASelected = true; // Default to calculated GIA
        public bool IsCalculatedGIASelected
        {
            get => _isCalculatedGIASelected;
            set
            {
                if (_isCalculatedGIASelected != value)
                {
                    _isCalculatedGIASelected = value;
                    OnPropertyChanged(nameof(IsCalculatedGIASelected));
                    // If we're selecting calculated GIA, ensure input GIA is deselected
                    if (value && _isInputGIASelected)
                    {
                        _isInputGIASelected = false;
                        OnPropertyChanged(nameof(IsInputGIASelected));
                    }
                    UpdateCarbonRating();
                }
            }
        }

        // Load volumes from the Revit model
        private void LoadVolumes()
        {
            // Get all materials in the document first
            var allMaterials = new FilteredElementCollector(_document)
                .OfClass(typeof(Material))
                .Cast<Material>()
                .ToDictionary(m => m.Name, m => m.Id);

            // Log available materials for debugging
            System.Diagnostics.Debug.WriteLine($"Found {allMaterials.Count} materials in document:");
            foreach (var mat in allMaterials.Keys.Take(10)) // Just show first 10
            {
                System.Diagnostics.Debug.WriteLine($"- {mat}");
            }


            // The volume finding code is buggy 
            // Find the closest matching materials
            ElementId concreteMatId = FindBestMaterialMatch(allMaterials, "Concrete", "WAL_Concrete_RC");
            ElementId steelMatId = FindBestMaterialMatch(allMaterials, "Steel", "WAL_Steel", "Metal");
            ElementId timberMatId = FindBestMaterialMatch(allMaterials, "Wood", "Wood - Dimensional Lumber", "Lumber");
            ElementId blockMatId = FindBestMaterialMatch(allMaterials, "Block", "WAL_Block", "Masonry");
            ElementId brickMatId = FindBestMaterialMatch(allMaterials, "Brick", "WAL_Brick");

            // Dictionary to store volumes by element type
            var concreteVolumes = new Dictionary<string, double>();
            var steelVolumes = new Dictionary<string, double>();
            var timberVolumes = new Dictionary<string, double>();
            var masonryVolumes = new Dictionary<string, double>();

            // Get all structural elements
            var structuralElements = GetAllStructuralElements(_document);

            // Process each element
            foreach (var element in structuralElements)
            {
                // Skip elements without materials
                var materials = element.GetMaterialIds(false);
                if (materials == null || materials.Count == 0)
                    continue;

                string elementType = GetElementType(element);
                if (string.IsNullOrEmpty(elementType))
                    continue;

                // Get material volumes
                foreach (ElementId materialId in materials)
                {
                    double volume = 0;
                    try
                    {
                        volume = element.GetMaterialVolume(materialId);
                    }
                    catch
                    {
                        continue; // Skip if we can't get volume
                    }

                    if (volume <= 0)
                        continue;

                    if (materialId.Equals(concreteMatId))
                    {
                        if (!concreteVolumes.ContainsKey(elementType))
                            concreteVolumes[elementType] = 0;
                        concreteVolumes[elementType] += volume;
                    }
                    else if (materialId.Equals(steelMatId))
                    {
                        if (!steelVolumes.ContainsKey(elementType))
                            steelVolumes[elementType] = 0;
                        steelVolumes[elementType] += volume;
                    }
                    else if (materialId.Equals(timberMatId))
                    {
                        if (!timberVolumes.ContainsKey(elementType))
                            timberVolumes[elementType] = 0;
                        timberVolumes[elementType] += volume;
                    }
                    else if (materialId.Equals(blockMatId) || materialId.Equals(brickMatId))
                    {
                        if (!masonryVolumes.ContainsKey(elementType))
                            masonryVolumes[elementType] = 0;
                        masonryVolumes[elementType] += volume;
                    }
                }
            }

            // Update the dictionary values
            foreach (var key in _items.Keys)
            {
                _items[key].ConcreteVolume = concreteVolumes.ContainsKey(key) ? concreteVolumes[key] / 35.315 : 0;
                _items[key].SteelVolume = steelVolumes.ContainsKey(key) ? steelVolumes[key] / 35.315 : 0;
                _items[key].TimberVolume = timberVolumes.ContainsKey(key) ? timberVolumes[key] / 35.315 : 0;
                _items[key].MasonryVolume = masonryVolumes.ContainsKey(key) ? masonryVolumes[key] / 35.315 : 0;
            }

            UpdateTotalCarbonEmbodiment();
            UpdateItemsView();
        }

        // Helper method to find the best matching material
        private ElementId FindBestMaterialMatch(Dictionary<string, ElementId> materials, params string[] possibleNames)
        {
            // First try exact matches
            foreach (var name in possibleNames)
            {
                if (materials.ContainsKey(name))
                    return materials[name];
            }

            // Then try contains matches (case insensitive)
            foreach (var material in materials)
            {
                foreach (var name in possibleNames)
                {
                    if (material.Key.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                        return material.Value;
                }
            }

            // Return null if no match found
            return null;
        }

        // Get all structural elements from the document
        private List<Element> GetAllStructuralElements(Document doc)
        {
            var elements = new List<Element>();

            // Add structural framing (beams)
            elements.AddRange(new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .WhereElementIsNotElementType()
                .ToElements());

            // Add structural columns
            elements.AddRange(new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .WhereElementIsNotElementType()
                .ToElements());

            // Add floors
            elements.AddRange(new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Floors)
                .WhereElementIsNotElementType()
                .ToElements());

            // Add walls
            elements.AddRange(new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .ToElements());

            // Add structural foundations
            elements.AddRange(new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .WhereElementIsNotElementType()
                .ToElements());

            return elements;
        }

        // Determine the element type (Beam, Column, Floor, etc.)
        private string GetElementType(Element element)
        {
            if (element is FamilyInstance familyInstance)
            {
                var structType = familyInstance.StructuralType;
                if (structType == StructuralType.Beam)
                    return "Beam";
                if (structType == StructuralType.Column)
                    return "Column";
                if (structType == StructuralType.Footing)
                    return "Foundation";

                // Check for piling (may need additional checks based on family/type name)
                var typeName = element.Name;
                if (typeName.Contains("Pile") || typeName.Contains("Piling"))
                    return "Piling";
            }
            else if (element is Wall wall)
            {
                // Check if it's an upstand wall
                var comments = element.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsString() ?? "";
                if (comments.Contains("Upstand") || wall.Name.Contains("Upstand"))
                    return "Upstand";

                return "Wall";
            }
            else if (element is Floor)
            {
                return "Floor";
            }
            else if (element is FootPrintRoof)
            {
                return "Floor"; // Group roofs with floors for now
            }

            // For foundations
            if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFoundation)
                return "Foundation";

            return null;
        }

        // Load GIA calculations
        private void LoadGIA()
        {
            GIA = 5000; // Example fixed value
            GIACalculated = GetTotalFloorArea(_document);
        }

        // Get total floor area from the document
        private double GetTotalFloorArea(Document doc)
        {
            var floors = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Floors)
                .WhereElementIsNotElementType()
                .ToElements();

            return floors.Sum(floor => floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)?.AsDouble() ?? 0) / 10.764; // Convert to square meters
        }

        // Method to get material volumes
        private Dictionary<string, double> GetVolumeOfMaterialsInComponent(Document doc, List<string> materialNames, string component, string subtype = null)
        {
            var materialVolumes = new Dictionary<string, double>();

            foreach (var material in materialNames)
            {
                // Use the flexible matching method instead
                var matId = RevitHelper.GetMaterialIdByFlexibleName(doc, material);
                if (matId == null)
                {
                    // If material is not found, log it and continue
                    System.Diagnostics.Debug.WriteLine($"Material not found: {material}");
                    materialVolumes[material] = 0;
                    continue;
                }

                double volume = 0;
                List<Element> elements = new List<Element>();

                // The rest of your existing code for getting elements...

                // Optional: if you want to skip filtering by subtype for now to get any elements with the material
                // elements = GetAllElementsForComponent(doc, component);

                foreach (var element in elements)
                {
                    try
                    {
                        // Try to get the volume directly if possible
                        double elementVolume = element.GetMaterialVolume(matId);
                        volume += elementVolume;
                    }
                    catch
                    {
                        // If GetMaterialVolume fails, try to calculate it from parameters
                        try
                        {
                            // Try to calculate volume from parameters
                            double elementVolume = CalculateElementVolume(element);

                            // If we have a volume and the element has this material, count it
                            if (elementVolume > 0 && ElementHasMaterial(element, matId))
                            {
                                volume += elementVolume;
                            }
                        }
                        catch
                        {
                            // Couldn't calculate volume, skip this element
                            continue;
                        }
                    }
                }

                materialVolumes[material] = volume;
            }

            return materialVolumes;
        }

        // Helper method to check if an element has a specific material
        private bool ElementHasMaterial(Element element, ElementId materialId)
        {
            try
            {
                var materialIds = element.GetMaterialIds(false);
                return materialIds.Contains(materialId);
            }
            catch
            {
                return false;
            }
        }

        // Helper method to calculate volume from parameters
        private double CalculateElementVolume(Element element)
        {
            // Try to get volume parameter
            Parameter volumeParam = element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
            if (volumeParam != null && volumeParam.HasValue)
            {
                return volumeParam.AsDouble();
            }

            // For walls
            if (element is Wall wall)
            {
                // Get wall width (it's a property, not a parameter)
                double width = wall.Width;
                // Get wall length and height
                Parameter lengthParam = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                Parameter heightParam = element.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);

                if (width > 0 && lengthParam != null && heightParam != null && lengthParam.HasValue && heightParam.HasValue)
                {
                    return width * lengthParam.AsDouble() * heightParam.AsDouble();
                }
            }

            // For floors and other area elements
            Parameter areaParam = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
            if (areaParam != null && areaParam.HasValue)
            {
                // For floors, try to get the actual thickness
                if (element is Floor)
                {
                    double thickness = 0;
                    // Try to get type element and check if it has a thickness parameter
                    try
                    {
                        ElementId typeId = element.GetTypeId();
                        if (typeId != ElementId.InvalidElementId)
                        {
                            Element typeElem = element.Document.GetElement(typeId);
                            Parameter typeThicknessParam = typeElem.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                            if (typeThicknessParam != null && typeThicknessParam.HasValue)
                            {
                                thickness = typeThicknessParam.AsDouble();
                            }
                        }
                    }
                    catch
                    {
                        // Use a default thickness if we can't determine it
                        thickness = 0.3; // Default to 300mm or approx 1ft
                    }

                    if (thickness > 0)
                    {
                        return areaParam.AsDouble() * thickness;
                    }
                }

                // For other elements with area, use a default thickness as fallback
                return areaParam.AsDouble() * 0.3; // Default thickness as fallback
            }

            // For structural framing elements (beams, columns)
            if (element is FamilyInstance)
            {
                // Try to calculate from length and section properties
                Parameter lengthParam = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                if (lengthParam != null && lengthParam.HasValue)
                {
                    double length = lengthParam.AsDouble();

                    // Try to get cross-section area
                    Parameter sectionAreaParam = element.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_AREA);
                    if (sectionAreaParam != null && sectionAreaParam.HasValue)
                    {
                        return length * sectionAreaParam.AsDouble();
                    }
                }
            }

            return 0;
        }
        private List<Element> FilterElementsBySubtype(List<Element> elements, string subtype)
        {
            // For steel elements
            if (subtype == "Open Section" || subtype == "Closed Section" || subtype == "Plates")
            {
                return elements.Where(e =>
                {
                    var comments = e.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsString() ?? "";
                    if (subtype == "Open Section")
                        return comments.Contains("UB") || comments.Contains("UC") || comments.Contains("PFC") || comments.Contains("Angle");
                    else if (subtype == "Closed Section")
                        return comments.Contains("SHS") || comments.Contains("RHS") || comments.Contains("CHS");
                    else // Plates
                        return comments.Contains("PL") || comments.Contains("Plate");
                }).ToList();
            }

            // For timber elements
            if (subtype == "Softwood" || subtype == "Glulam" || subtype == "LVL" || subtype == "CLT")
            {
                return elements.Where(e =>
                {
                    var comments = e.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsString() ?? "";
                    return comments.Contains(subtype);
                }).ToList();
            }

            return elements; // Return all elements if no filtering is needed
        }

        private double GetElementMaterialVolume(Element element, ElementId materialId)
        {
            try
            {
                return element.GetMaterialVolume(materialId);
            }
            catch
            {
                return 0;
            }
        }

        // Update total carbon embodiment
        public void UpdateTotalCarbonEmbodiment()
        {
            TotalCarbonEmbodiment = _items.Values.Sum(item => item.SubTotalCarbon);
            UpdateCarbonRating();
        }

        // Update the CarbonRating calculation to use the selected GIA
        private void UpdateCarbonRating()
        {
            double selectedGIA = IsInputGIASelected ? GIA : GIACalculated;
            if (selectedGIA > 0)
            {
                CarbonRating = TotalCarbonEmbodiment / selectedGIA;
            }
            else
            {
                CarbonRating = 0;
            }
        }

        // Method to get project name
        private void GetProjectName()
        {
            if (_document != null)
            {
                var projectInfo = _document.ProjectInformation;
                ProjectName = projectInfo?.Name ?? "Unknown Project";
            }
        }

        private void GetProjectNumber()
        {
            if (_document != null)
            {
                var projectInfo = _document.ProjectInformation;
                ProjectNumber = projectInfo?.Number ?? "Unknown Project";
            }
        }

        private string GetDataFilePath()
        {
            string directory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData);
            string filePath = System.IO.Path.Combine(directory, "ECATPluginData.xml");
            return filePath;
        }

        public void SaveData()
        {
            var dataFilePath = GetDataFilePath();
            XDocument doc = new XDocument(
                new XElement("ECATPluginData",
                    new XElement("ProjectName", ProjectName),
                    new XElement("ProjectNumber", ProjectNumber),
                    new XElement("GIA", GIA),
                    new XElement("GIACalculated", GIACalculated),
                    new XElement("IsInputGIASelected", IsInputGIASelected),
                    new XElement("IsCalculatedGIASelected", IsCalculatedGIASelected),
                    new XElement("Items",
                        _items.Select(item =>
                            new XElement("Item",
                                new XAttribute("Name", item.Key),
                                new XElement("EC", item.Value.EC),
                                new XElement("RebarDensity", item.Value.RebarDensity),
                                new XElement("SteelType", item.Value.SteelType),
                                new XElement("SteelSource", item.Value.SteelSource),
                                new XElement("TimberType", item.Value.TimberType),
                                new XElement("TimberSource", item.Value.TimberSource),
                                new XElement("MasonryType", item.Value.MasonryType),
                                new XElement("Module", item.Value.Module)
                            )
                        )
                    )
                )
            );

            doc.Save(dataFilePath);
        }

        public void LoadData()
        {
            var dataFilePath = GetDataFilePath();
            if (System.IO.File.Exists(dataFilePath))
            {
                try
                {
                    XDocument doc = XDocument.Load(dataFilePath);
                    var root = doc.Element("ECATPluginData");

                    // Load project details
                    ProjectName = root.Element("ProjectName")?.Value ?? ProjectName;
                    ProjectNumber = root.Element("ProjectNumber")?.Value ?? ProjectNumber;

                    // Load GIA settings
                    if (root.Element("GIA") != null)
                        GIA = double.Parse(root.Element("GIA").Value);

                    if (root.Element("GIACalculated") != null)
                        GIACalculated = double.Parse(root.Element("GIACalculated").Value);

                    // Load GIA selection state
                    bool isInputGIASelected = false;
                    bool isCalculatedGIASelected = true;

                    if (root.Element("IsInputGIASelected") != null)
                        isInputGIASelected = bool.Parse(root.Element("IsInputGIASelected").Value);

                    if (root.Element("IsCalculatedGIASelected") != null)
                        isCalculatedGIASelected = bool.Parse(root.Element("IsCalculatedGIASelected").Value);

                    // Set the GIA selection state after loading both values
                    IsInputGIASelected = isInputGIASelected;
                    IsCalculatedGIASelected = isCalculatedGIASelected;

                    // Load item data
                    foreach (var itemElement in root.Element("Items").Elements("Item"))
                    {
                        string name = itemElement.Attribute("Name")?.Value;
                        if (_items.ContainsKey(name))
                        {
                            var itemData = _items[name];

                            // Load carbon values
                            if (itemElement.Element("EC") != null)
                                itemData.EC = double.Parse(itemElement.Element("EC").Value);

                            if (itemElement.Element("RebarDensity") != null)
                                itemData.RebarDensity = double.Parse(itemElement.Element("RebarDensity").Value);

                            // Load material type properties
                            if (itemElement.Element("SteelType") != null)
                                itemData.SteelType = itemElement.Element("SteelType").Value;

                            if (itemElement.Element("SteelSource") != null)
                                itemData.SteelSource = itemElement.Element("SteelSource").Value;

                            if (itemElement.Element("TimberType") != null)
                                itemData.TimberType = itemElement.Element("TimberType").Value;

                            if (itemElement.Element("TimberSource") != null)
                                itemData.TimberSource = itemElement.Element("TimberSource").Value;

                            if (itemElement.Element("MasonryType") != null)
                                itemData.MasonryType = itemElement.Element("MasonryType").Value;

                            if (itemElement.Element("Module") != null)
                                itemData.Module = itemElement.Element("Module").Value;
                        }
                    }

                    // Update calculations after loading data
                    UpdateTotalCarbonEmbodiment();
                    UpdateItemsView(); // Update ItemsView after loading data
                }
                catch (Exception ex)
                {
                    // Log error but continue - don't crash on corrupted save file
                    System.Diagnostics.Debug.WriteLine($"Error loading saved data: {ex.Message}");
                }
            }
        }

        // Get available steel types for UI binding
        public List<string> AvailableSteelTypes
        {
            get { return _steelData.Keys.ToList(); }
        }

        // Get available steel sources for UI binding
        // Get available steel sources for a specific steel type
        public List<string> GetAvailableSteelSources(string steelType)
        {
            if (_steelData.ContainsKey(steelType))
            {
                return _steelData[steelType].Keys.ToList();
            }
            return new List<string>();
        }

        

        // Get available timber types for UI binding
        public List<string> AvailableTimberTypes
        {
            get { return _timberData.Keys.ToList(); }
        }

        // Get available timber sources for a specific timber type
        public List<string> GetAvailableTimberSources(string timberType)
        {
            if (_timberData.ContainsKey(timberType))
            {
                return _timberData[timberType].Keys.ToList();
            }
            return new List<string>();
        }

        // Get available masonry types for UI binding
        public List<string> AvailableMasonryTypes
        {
            get { return _masonryData.Keys.ToList(); }
        }

        // Get available modules for UI binding
        public List<string> AvailableModules
        {
            get { return new List<string> { "A1-A3", "A4", "A5" }; }
        }

        public List<string> AvailableSteelSources
        {
            get
            {
                var sources = new List<string>();
                foreach (var type in _steelData.Keys)
                {
                    sources.AddRange(_steelData[type].Keys);
                }
                return sources.Distinct().ToList();
            }
        }

        // List of all possible timber sources (union of all timber type sources)
        public List<string> AvailableTimberSources
        {
            get
            {
                var sources = new List<string>();
                foreach (var type in _timberData.Keys)
                {
                    sources.AddRange(_timberData[type].Keys);
                }
                return sources.Distinct().ToList();
            }
        }

    }
}