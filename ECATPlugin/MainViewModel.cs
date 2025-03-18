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
using System.Windows;

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
            try
            {
                // Create an instance of RatingPopup with the CarbonRating value
                RatingPopup ratingPopup = new RatingPopup(CarbonRating, this);

                // Show the window as non-modal
                ratingPopup.Show();
            }
            catch (Exception ex)
            {
                // Handle any exceptions to prevent Revit from crashing
                System.Diagnostics.Debug.WriteLine($"Error showing ratings: {ex.Message}");
                MessageBox.Show($"Unable to show ratings: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
                    OnPropertyChanged(nameof(ConcreteCarbon));
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
                    OnPropertyChanged(nameof(ConcreteCarbon));
                    OnPropertyChanged(nameof(SubTotalCarbon));
                }
            }

            private string _steelType = "Open Section"; // Open Section, Closed Section, Plates
            public string SteelType
            {
                get => _steelType;
                set
                {
                    if (_steelType != value)
                    {
                        _steelType = value;
                        OnPropertyChanged(nameof(SteelType));
                        OnPropertyChanged(nameof(SteelCarbon));
                        OnPropertyChanged(nameof(SubTotalCarbon));
                    }
                }
            }

            private string _steelSource = "UK"; // UK, Global, UK (Reused)
            public string SteelSource
            {
                get => _steelSource;
                set
                {
                    if (_steelSource != value)
                    {
                        _steelSource = value;
                        OnPropertyChanged(nameof(SteelSource));
                        OnPropertyChanged(nameof(SteelCarbon));
                        OnPropertyChanged(nameof(SubTotalCarbon));
                    }
                }
            }

            // Timber specific properties
            private string _timberType = "Softwood"; // Softwood, Glulam, LVL, CLT
            public string TimberType
            {
                get => _timberType;
                set
                {
                    if (_timberType != value)
                    {
                        _timberType = value;
                        OnPropertyChanged(nameof(TimberType));
                        OnPropertyChanged(nameof(TimberCarbon));
                        OnPropertyChanged(nameof(SubTotalCarbon));
                    }
                }
            }

            private string _timberSource = "Global"; // UK & EU, Global
            public string TimberSource
            {
                get => _timberSource;
                set
                {
                    if (_timberSource != value)
                    {
                        _timberSource = value;
                        OnPropertyChanged(nameof(TimberSource));
                        OnPropertyChanged(nameof(TimberCarbon));
                        OnPropertyChanged(nameof(SubTotalCarbon));
                    }
                }
            }

            // Masonry specific properties
            private string _masonryType = "Blockwork"; // Blockwork, Brickwork
            public string MasonryType
            {
                get => _masonryType;
                set
                {
                    if (_masonryType != value)
                    {
                        _masonryType = value;
                        OnPropertyChanged(nameof(MasonryType));
                        OnPropertyChanged(nameof(MasonryCarbon));
                        OnPropertyChanged(nameof(SubTotalCarbon));
                    }
                }
            }

    

            // Module properties for all materials
            private string _module = "A1-A3"; // A1-A3, A4, A5
            public string Module
            {
                get => _module;
                set
                {
                    if (_module != value)
                    {
                        _module = value;
                        OnPropertyChanged(nameof(Module));
                        // Notify all carbon properties that might be affected
                        OnPropertyChanged(nameof(SteelCarbon));
                        OnPropertyChanged(nameof(TimberCarbon));
                        OnPropertyChanged(nameof(MasonryCarbon));
                        OnPropertyChanged(nameof(SubTotalCarbon));
                    }
                }
            }

            public double ConcreteCarbon => CalculateConcreteCarbon();
            public double SteelCarbon => CalculateSteelCarbon();
            public double TimberCarbon => CalculateTimberCarbon();
            public double MasonryCarbon => CalculateMasonryCarbon();

            // Modify this property to use the individual calculations
            public double SubTotalCarbon => ConcreteCarbon + SteelCarbon + TimberCarbon + MasonryCarbon;

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

            // Add this new method for concrete carbon calculation
            private double CalculateConcreteCarbon()
            {
                return (ConcreteVolume * EC + ConcreteVolume * RebarDensity) * 0.785;
            }

            // Keep these methods as they are
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

        private string _phase;
        public string Phase
        {
            get => _phase;
            set
            {
                _phase = value;
                OnPropertyChanged(nameof(Phase));
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

                    // Concrete is applicable for all structural elements
                    if (materialId.Equals(concreteMatId))
                    {
                        if (!concreteVolumes.ContainsKey(elementType))
                            concreteVolumes[elementType] = 0;
                        concreteVolumes[elementType] += volume;
                    }

                    // Steel is only for beams, columns, floors, foundations, and pilings
                    else if (materialId.Equals(steelMatId) && IsItemValidForMaterial(elementType, "Steel"))
                    {
                        if (!steelVolumes.ContainsKey(elementType))
                            steelVolumes[elementType] = 0;
                        steelVolumes[elementType] += volume;
                    }

                    // Timber is only for beams, columns, walls, and floors
                    else if (materialId.Equals(timberMatId) && IsItemValidForMaterial(elementType, "Timber"))
                    {
                        if (!timberVolumes.ContainsKey(elementType))
                            timberVolumes[elementType] = 0;
                        timberVolumes[elementType] += volume;
                    }

                    // Masonry is only for walls
                    else if ((materialId.Equals(blockMatId) || materialId.Equals(brickMatId)) &&
                             IsItemValidForMaterial(elementType, "Masonry"))
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

                // Only set steel volumes for compatible component types
                if (IsItemValidForMaterial(key, "Steel"))
                    _items[key].SteelVolume = steelVolumes.ContainsKey(key) ? steelVolumes[key] / 35.315 : 0;
                else
                    _items[key].SteelVolume = 0;

                // Only set timber volumes for compatible component types
                if (IsItemValidForMaterial(key, "Timber"))
                    _items[key].TimberVolume = timberVolumes.ContainsKey(key) ? timberVolumes[key] / 35.315 : 0;
                else
                    _items[key].TimberVolume = 0;

                // Only set masonry volumes for compatible component types
                if (IsItemValidForMaterial(key, "Masonry"))
                    _items[key].MasonryVolume = masonryVolumes.ContainsKey(key) ? masonryVolumes[key] / 35.315 : 0;
                else
                    _items[key].MasonryVolume = 0;
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
        // Update the TotalCarbonEmbodiment calculation in MainViewModel.cs
        public void UpdateTotalCarbonEmbodiment()
        {
            // Calculate the sum of the subtotal carbon values for all materials across all items
            TotalCarbonEmbodiment = _items.Values.Sum(item => item.SubTotalCarbon) / 1000; // Convert from kg to tonnes

            // Update the carbon rating based on the selected GIA
            UpdateCarbonRating();
        }

        // Update the CarbonRating calculation to use the selected GIA
        private void UpdateCarbonRating()
        {
            double selectedGIA = IsInputGIASelected ? GIA : GIACalculated;
            if (selectedGIA > 0)
            {
                // Calculate kgCO2e/m² by converting total carbon back to kg and dividing by GIA
                CarbonRating = (TotalCarbonEmbodiment * 1000) / selectedGIA;
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

        // Update these methods in MainViewModel.cs

        private string GetDataFilePath()
        {
            string directory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData);
            string projectSubfolder = "ECATPluginData";

            // Create the base directory if it doesn't exist
            string fullDirectory = System.IO.Path.Combine(directory, projectSubfolder);
            if (!System.IO.Directory.Exists(fullDirectory))
            {
                System.IO.Directory.CreateDirectory(fullDirectory);
            }

            // Get a unique identifier for the current project
            string projectIdentifier = GetProjectIdentifier();

            // Create a filename using the project identifier
            string fileName = $"ECAT_{projectIdentifier}.xml";

            string filePath = System.IO.Path.Combine(fullDirectory, fileName);
            return filePath;
        }

        // Helper method to get a unique identifier for the current project
        private string GetProjectIdentifier()
        {
            // First try to use project number and cleaned project name
            if (!string.IsNullOrEmpty(ProjectNumber) && !string.IsNullOrEmpty(ProjectName))
            {
                // Create a sanitized identifier
                string identifier = $"{ProjectNumber}_{CleanFileName(ProjectName)}";
                return identifier;
            }

            // If project number/name aren't available or are empty, use the document's unique ID
            if (_document != null)
            {
                // Use the document's UniqueId as a fallback
                string docId = _document.ProjectInformation?.UniqueId ?? _document.Title;
                if (!string.IsNullOrEmpty(docId))
                {
                    return CleanFileName(docId);
                }
            }

            // Last resort: use the document title or a timestamp if nothing else is available
            return CleanFileName(_document?.Title ?? DateTime.Now.ToString("yyyyMMdd_HHmmss"));
        }

        // Helper method to clean a string for use in a filename
        private string CleanFileName(string input)
        {
            if (string.IsNullOrEmpty(input))
                return "Unknown";

            // Replace invalid filename characters with underscores
            char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
            string result = new string(input.Select(c => invalidChars.Contains(c) ? '_' : c).ToArray());

            // Ensure it's not too long
            if (result.Length > 50)
                result = result.Substring(0, 50);

            return result;
        }

        public void SaveData()
        {
            try
            {
                var dataFilePath = GetDataFilePath();

                XDocument doc = new XDocument(
                    new XElement("ECATPluginData",
                        new XElement("ProjectName", ProjectName),
                        new XElement("ProjectNumber", ProjectNumber),
                        new XElement("Phase", Phase),
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
                                    new XElement("ConcreteVolume", item.Value.ConcreteVolume),
                                    new XElement("SteelVolume", item.Value.SteelVolume),
                                    new XElement("TimberVolume", item.Value.TimberVolume),
                                    new XElement("MasonryVolume", item.Value.MasonryVolume),
                                    new XElement("SteelType", item.Value.SteelType),
                                    new XElement("SteelSource", item.Value.SteelSource),
                                    new XElement("TimberType", item.Value.TimberType),
                                    new XElement("TimberSource", item.Value.TimberSource),
                                    new XElement("MasonryType", item.Value.MasonryType),
                                    new XElement("Module", item.Value.Module)
                                )
                            )
                        ),
                        new XElement("SavedDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                    )
                );

                doc.Save(dataFilePath);
                System.Diagnostics.Debug.WriteLine($"Data saved to {dataFilePath}");
            }
            catch (Exception ex)
            {
                // Log the error but don't crash
                System.Diagnostics.Debug.WriteLine($"Error saving data: {ex.Message}");
            }
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

                    if (root.Element("Phase") != null)
                        Phase = root.Element("Phase")?.Value ?? "";

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
                    var itemsElement = root.Element("Items");
                    if (itemsElement != null)
                    {
                        foreach (var itemElement in itemsElement.Elements("Item"))
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

                                // Load volume values - new additions
                                if (itemElement.Element("ConcreteVolume") != null)
                                    itemData.ConcreteVolume = double.Parse(itemElement.Element("ConcreteVolume").Value);

                                if (itemElement.Element("SteelVolume") != null)
                                    itemData.SteelVolume = double.Parse(itemElement.Element("SteelVolume").Value);

                                if (itemElement.Element("TimberVolume") != null)
                                    itemData.TimberVolume = double.Parse(itemElement.Element("TimberVolume").Value);

                                if (itemElement.Element("MasonryVolume") != null)
                                    itemData.MasonryVolume = double.Parse(itemElement.Element("MasonryVolume").Value);

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
                    }

                    // Update calculations after loading data
                    UpdateTotalCarbonEmbodiment();
                    UpdateItemsView(); // Update ItemsView after loading data

                    System.Diagnostics.Debug.WriteLine($"Data loaded from {dataFilePath}");
                }
                catch (Exception ex)
                {
                    // Log error but continue - don't crash on corrupted save file
                    System.Diagnostics.Debug.WriteLine($"Error loading saved data: {ex.Message}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"No saved data found at {dataFilePath}");
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

        // New properties to expose filtered collections
        private ObservableCollection<KeyValuePair<string, ItemData>> _steelItemsView;
        public ObservableCollection<KeyValuePair<string, ItemData>> SteelItemsView
        {
            get
            {
                if (_steelItemsView == null)
                {
                    _steelItemsView = new ObservableCollection<KeyValuePair<string, ItemData>>(
                        _items.Where(item =>
                            item.Key == "Beam" ||
                            item.Key == "Column" ||
                            item.Key == "Floor" ||
                            item.Key == "Foundation" ||
                            item.Key == "Piling")
                    );
                }
                return _steelItemsView;
            }
        }

        private ObservableCollection<KeyValuePair<string, ItemData>> _timberItemsView;
        public ObservableCollection<KeyValuePair<string, ItemData>> TimberItemsView
        {
            get
            {
                if (_timberItemsView == null)
                {
                    _timberItemsView = new ObservableCollection<KeyValuePair<string, ItemData>>(
                        _items.Where(item =>
                            item.Key == "Beam" ||
                            item.Key == "Column" ||
                            item.Key == "Wall" ||
                            item.Key == "Floor")
                    );
                }
                return _timberItemsView;
            }
        }

        private ObservableCollection<KeyValuePair<string, ItemData>> _masonryItemsView;
        public ObservableCollection<KeyValuePair<string, ItemData>> MasonryItemsView
        {
            get
            {
                if (_masonryItemsView == null)
                {
                    _masonryItemsView = new ObservableCollection<KeyValuePair<string, ItemData>>(
                        _items.Where(item => item.Key == "Wall")
                    );
                }
                return _masonryItemsView;
            }
        }

        // Function to update all filtered views when the main collection changes
        private void UpdateFilteredItemsViews()
        {
            _steelItemsView = null;
            OnPropertyChanged(nameof(SteelItemsView));

            _timberItemsView = null;
            OnPropertyChanged(nameof(TimberItemsView));

            _masonryItemsView = null;
            OnPropertyChanged(nameof(MasonryItemsView));
        }

        // Modify the existing UpdateItemsView method to also update filtered views
        private void UpdateItemsView()
        {
            _itemsView = null;
            OnPropertyChanged(nameof(ItemsView));

            // Also update the filtered views
            UpdateFilteredItemsViews();
        }

        // Helper method to check if an item should have non-zero volumes based on material type
        public bool ShouldShowItemForMaterial(string itemKey, string materialType)
        {
            switch (materialType)
            {
                case "Steel":
                    return itemKey == "Beam" ||
                           itemKey == "Column" ||
                           itemKey == "Floor" ||
                           itemKey == "Foundation" ||
                           itemKey == "Piling";
                case "Timber":
                    return itemKey == "Beam" ||
                           itemKey == "Column" ||
                           itemKey == "Wall" ||
                           itemKey == "Floor";
                case "Masonry":
                    return itemKey == "Wall";
                default:
                    return true;
            }
        }

        // Filter an item's volume based on material compatibility
        public bool IsItemValidForMaterial(string itemKey, string materialType)
        {
            bool isValid = ShouldShowItemForMaterial(itemKey, materialType);

            // Ensure volumes are zero for incompatible items
            if (!isValid && _items.ContainsKey(itemKey))
            {
                switch (materialType)
                {
                    case "Steel":
                        if (_items[itemKey].SteelVolume > 0)
                            _items[itemKey].SteelVolume = 0;
                        break;
                    case "Timber":
                        if (_items[itemKey].TimberVolume > 0)
                            _items[itemKey].TimberVolume = 0;
                        break;
                    case "Masonry":
                        if (_items[itemKey].MasonryVolume > 0)
                            _items[itemKey].MasonryVolume = 0;
                        break;
                }
            }

            return isValid;
        }
    

    }
}