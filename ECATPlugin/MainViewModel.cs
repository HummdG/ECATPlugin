using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ECATPlugin.Helpers;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows;
using System.Xml.Linq;
using System;



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
using Autodesk.Revit.UI;

namespace ECATPlugin
{
    public class MainViewModel : INotifyPropertyChanged
    {
        // Public method to trigger property change notifications
        public void TriggerPropertyChanged(string propertyName)
        {
            OnPropertyChanged(propertyName);
        }

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

            public void SetTimberECFactorFromSaved(double value, bool isManual)
            {
                _timberECFactor = value;
                _isManualTimberECFactor = isManual;
                OnPropertyChanged(nameof(TimberECFactor));
                OnPropertyChanged(nameof(TimberCarbon));
                OnPropertyChanged(nameof(SubTotalCarbon));
            }
            public void NotifyPropertyChanged(string propertyName)
            {
                OnPropertyChanged(propertyName);
            }

            public void ForcePropertyChange(string propertyName)
            {
                OnPropertyChanged(propertyName);
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
                    OnPropertyChanged(nameof(SteelCarbon)); // Add notification for SteelCarbon
                    OnPropertyChanged(nameof(SubTotalCarbon)); // Also update total carbon
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

            private double _timberECFactor;
            public double TimberECFactor
            {
                get => _timberECFactor;
                set
                {
                    _timberECFactor = value;
                    OnPropertyChanged(nameof(TimberECFactor));
                    OnPropertyChanged(nameof(TimberCarbon));
                    OnPropertyChanged(nameof(SubTotalCarbon));

                    // Add this line to flag that this was manually set
                    _isManualTimberECFactor = true;
                }
            }

            // 2. Add a new property to track if the TimberECFactor was manually set
            private bool _isManualTimberECFactor = false;
            public bool IsManualTimberECFactor
            {
                get => _isManualTimberECFactor;
                set
                {
                    _isManualTimberECFactor = value;
                    OnPropertyChanged(nameof(IsManualTimberECFactor));
                }
            }

            private double _masonryECFactor;
            public double MasonryECFactor
            {
                get => _masonryECFactor;
                set
                {
                    _masonryECFactor = value;
                    OnPropertyChanged(nameof(MasonryECFactor));
                    OnPropertyChanged(nameof(MasonryCarbon));
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

                        // This is already correct, but making sure we trigger UpdateTotalCarbonEmbodiment
                        System.Diagnostics.Debug.WriteLine($"Steel source changed to {value}, updating carbon calculations");
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

            // 4. Modify GetTimberECFactor method to respect manually set values
            private double GetTimberECFactor()
            {
                // If manual value has been set, use it
                if (_isManualTimberECFactor && _timberECFactor > 0)
                    return _timberECFactor;

                // Otherwise calculate from defaults
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

                // Use manual value if set
                if (_masonryECFactor > 0)
                {
                    ecFactor = _masonryECFactor;
                }
                else
                {
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
        private Dictionary<string, string> _steelSectionMapping;
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



        // Fixed InitializeMaterialData method
        private void InitializeItems()
        {
            _items["Beams"] = new ItemData { EC = 250, RebarDensity = 225, TimberECFactor = 0.263 };
            _items["Walls"] = new ItemData { EC = 250, RebarDensity = 180, TimberECFactor = 0.263 };
            _items["Upstands"] = new ItemData { EC = 250, RebarDensity = 180, TimberECFactor = 0.263 };
            _items["Columns"] = new ItemData { EC = 250, RebarDensity = 250, TimberECFactor = 0.263 };
            _items["Floors"] = new ItemData { EC = 215, RebarDensity = 120, TimberECFactor = 0.263 };
            _items["Foundations"] = new ItemData { EC = 215, RebarDensity = 110, TimberECFactor = 0.263 };
            _items["Pilings"] = new ItemData { EC = 140, RebarDensity = 75, TimberECFactor = 0.263 };

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

            _steelSectionMapping = new Dictionary<string, string>
            {
                // Open Sections
                { "ASB", "Open Sections" },
                { "PFC", "Open Sections" },
                { "RSJ", "Open Sections" },
                { "SFB", "Open Sections" },
                { "TUB", "Open Sections" },
                { "TUC", "Open Sections" },
                { "UB", "Open Sections" },
                { "UBP", "Open Sections" },
                { "UC", "Open Sections" },
    
                // Closed Sections
                { "SHS", "Closed Sections" },
                { "CHS", "Closed Sections" },
                { "RHS", "Closed Sections" },

                { "PLT", "Plates" },
                { "RSA", "Plates" }

            };
        }
        // Helper method to determine section type
        private string DetermineSectionType(Element element)
        {
            // Get family and type name
            string familyName = "";
            string typeName = "";

            if (element is FamilyInstance familyInstance)
            {
                try
                {
                    ElementId typeId = element.GetTypeId();
                    if (typeId != ElementId.InvalidElementId)
                    {
                        ElementType type = element.Document.GetElement(typeId) as ElementType;
                        if (type != null)
                        {
                            familyName = type.FamilyName ?? "";
                            typeName = type.Name ?? "";
                        }
                    }
                }
                catch { }
            }

            // First, check for any WAL_SFRA_STL family patterns which are the most reliable
            if (familyName.Contains("WAL_SFRA_STL"))
            {
                // Check for section patterns in the format WAL_SFRA_STL-XXX where XXX is the section identifier
                if (familyName.Contains("-"))
                {
                    string[] parts = familyName.Split('-');
                    if (parts.Length > 1)
                    {
                        string suffix = parts[1];

                        // Extract the alphabetic prefix (e.g., TUB from TUB191x229x34)
                        string prefix = new string(suffix.TakeWhile(char.IsLetter).ToArray());

                        // Check if it's an Open Section
                        string[] openSectionIds = { "ASB", "PFC", "RSJ", "SFB", "TUB", "TUC", "UB", "UBP", "UC" };
                        if (openSectionIds.Contains(prefix, StringComparer.OrdinalIgnoreCase))
                        {
                            return "Open Sections";
                        }

                        // Check if it's a Closed Section
                        string[] closedSectionIds = { "SHS", "CHS", "RHS" };
                        if (closedSectionIds.Contains(prefix, StringComparer.OrdinalIgnoreCase))
                        {
                            return "Closed Sections";
                        }

                        // Check for explicit plate identifiers
                        if (prefix.Equals("PLT", StringComparison.OrdinalIgnoreCase) ||
                            prefix.Equals("RSA", StringComparison.OrdinalIgnoreCase))
                        {
                            return "Plates";
                        }

                        // If we got here with a WAL_SFRA_STL element but couldn't identify the section type,
                        // still return Plates as this is definitely a steel element
                        return "Plates";
                    }
                }
            }

            // If not a WAL_SFRA_STL element, look for section identifiers in the names

            // Check for Open Section identifiers
            string[] openSectionKeywords = { "ASB", "PFC", "RSJ", "SFB", "TUB", "TUC", "UB", "UBP", "UC" };
            foreach (var keyword in openSectionKeywords)
            {
                // Use case-insensitive comparison
                if (familyName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    typeName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "Open Sections";
                }
            }

            // Check for Closed Section identifiers
            string[] closedSectionKeywords = { "SHS", "CHS", "RHS" };
            foreach (var keyword in closedSectionKeywords)
            {
                // Use case-insensitive comparison
                if (familyName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    typeName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "Closed Sections";
                }
            }

            // Check for Plate identifiers
            string[] plateKeywords = { "PLT", "PLATE", "RSA", "ANGLE" };
            foreach (var keyword in plateKeywords)
            {
                // Use case-insensitive comparison
                if (familyName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    typeName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "Plates";
                }
            }

            // If we reached this point, we couldn't identify the section type.
            // Return an empty string to indicate this is not a recognizable steel element.
            return "";
        }
        private void ItemData_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // Update TotalCarbonEmbodiment if relevant properties change
            if (e.PropertyName == nameof(ItemData.Volume) ||
                e.PropertyName == nameof(ItemData.EC) ||
                e.PropertyName == nameof(ItemData.RebarDensity) ||
                e.PropertyName == nameof(ItemData.SteelCarbon) || // Add SteelCarbon
                e.PropertyName == nameof(ItemData.SteelVolume) || // Add SteelVolume 
                e.PropertyName == nameof(ItemData.SteelType) ||   // Add SteelType
                e.PropertyName == nameof(ItemData.SteelSource) || // Add SteelSource
                e.PropertyName == nameof(ItemData.SubTotalCarbon))
            {
                System.Diagnostics.Debug.WriteLine($"ItemData property changed: {e.PropertyName}, updating total carbon embodiment");
                UpdateTotalCarbonEmbodiment();
            }
        }

        public void Dispose()
        {
            try
            {
                // First, save the current data
                SaveData();

                // Then remove event handlers from items
                foreach (var item in _items.Values)
                {
                    try
                    {
                        item.PropertyChanged -= ItemData_PropertyChanged;
                    }
                    catch (Exception ex)
                    {
                        // Log but continue with other items
                        System.Diagnostics.Debug.WriteLine($"Error removing event handler: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the error but don't throw during disposal
                System.Diagnostics.Debug.WriteLine($"Error during MainViewModel disposal: {ex.Message}");
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

        private Dictionary<string, Dictionary<string, double>> _steelVolumeBySection;

        // Load volumes from the Revit model
        private void LoadVolumes()
        {
            // Get all materials in the document first
            var allMaterials = new FilteredElementCollector(_document)
                .OfClass(typeof(Material))
                .Cast<Material>()
                .ToDictionary(m => m.Name, m => m.Id);

            // Find the closest matching materials
            ElementId concreteMatId = FindBestMaterialMatch(allMaterials, "Concrete", "WAL_Concrete_RC");
            ElementId steelMatId = FindBestMaterialMatch(allMaterials, "Steel", "WAL_Steel", "Metal");
            ElementId timberMatId = FindBestMaterialMatch(allMaterials, "Wood", "Wood - Dimensional Lumber", "Lumber");
            ElementId blockMatId = FindBestMaterialMatch(allMaterials, "Block", "WAL_Block", "Masonry");
            ElementId brickMatId = FindBestMaterialMatch(allMaterials, "Brick", "WAL_Brick");


            // Initialize the steel volume by section dictionary
            _steelVolumeBySection = new Dictionary<string, Dictionary<string, double>>();
            foreach (var key in _items.Keys)
            {
                if (IsItemValidForMaterial(key, "Steel"))
                {
                    _steelVolumeBySection[key] = new Dictionary<string, double>
            {
                { "Open Sections", 0 },
                { "Closed Sections", 0 },
                { "Plates", 0 }
            };
                }
            }

            // Get all structural elements
            var structuralElements = GetAllStructuralElements(_document);

            // Dictionary to store volumes by element type
            var concreteVolumes = new Dictionary<string, double>();
            var steelVolumes = new Dictionary<string, double>();
            var timberVolumes = new Dictionary<string, double>();
            var masonryVolumes = new Dictionary<string, double>();

            // Key change: keep a separate dictionary just for true steel elements' volumes by category
            var steelVolumesByCategory = new Dictionary<string, Dictionary<string, double>>();
            foreach (var key in _items.Keys)
            {
                if (IsItemValidForMaterial(key, "Steel"))
                {
                    steelVolumesByCategory[key] = new Dictionary<string, double>
            {
                { "Open Sections", 0 },
                { "Closed Sections", 0 },
                { "Plates", 0 }
            };
                }
            }

            // Process each element
            foreach (var element in structuralElements)
            {
                // Check for Phasing parameter - only process elements in the "Proposed" phase
                Parameter phasingParam = element.LookupParameter("Phasing");
                if (phasingParam != null && phasingParam.HasValue)
                {
                    string phasing = phasingParam.AsString();
                    if (phasing == null || !phasing.Equals("Proposed", StringComparison.OrdinalIgnoreCase))
                    {
                        continue; // Skip if not "Proposed"
                    }
                }

                // Skip elements without materials
                var materials = element.GetMaterialIds(false);
                if (materials == null || materials.Count == 0)
                    continue;

                string elementType = GetElementType(element);
                if (string.IsNullOrEmpty(elementType))
                    continue;

                // Process each material in the element
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
                        // Add to regular steel volumes as before
                        if (!steelVolumes.ContainsKey(elementType))
                            steelVolumes[elementType] = 0;
                        steelVolumes[elementType] += volume;

                        // Determine steel section type ONLY for elements with steel material ID
                        if (steelVolumesByCategory.ContainsKey(elementType))
                        {
                            // Get the section type for this steel element
                            string sectionType = DetermineSectionType(element);

                            // Only add volume if the section type was identified
                            if (!string.IsNullOrEmpty(sectionType))
                            {
                                // Add the volume to the appropriate section category
                                steelVolumesByCategory[elementType][sectionType] += volume;
                            }
                            else
                            {
                                // Skip this element for section categorization since it's not a recognizable steel element
                                // This prevents non-steel elements from being categorized as "Plates"
                            }
                        }
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

            // Copy the steel volumes by category to the field for use elsewhere
            _steelVolumeBySection = steelVolumesByCategory;

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
            _steelHierarchy = null;
            OnPropertyChanged(nameof(SteelHierarchy));
        }


        // Helper method to find the best matching material
        private ElementId FindBestMaterialMatch(Dictionary<string, ElementId> materials, params string[] possibleNames)
        {
            System.Diagnostics.Debug.WriteLine($"Searching for material matching: {string.Join(", ", possibleNames)}");

            // First try exact matches
            foreach (var name in possibleNames)
            {
                if (materials.ContainsKey(name))
                {
                    System.Diagnostics.Debug.WriteLine($"Found exact match: {name}");
                    return materials[name];
                }
            }

            // Then try contains matches (case insensitive)
            foreach (var material in materials)
            {
                foreach (var name in possibleNames)
                {
                    if (material.Key.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"Found partial match: {material.Key} for {name}");
                        return material.Value;
                    }
                }
            }

            // Try additional materials based on color or other properties
            if (possibleNames.Contains("Steel") || possibleNames.Contains("WAL_Steel") || possibleNames.Contains("Metal"))
            {
                // Look for any material with "metal" in the name
                foreach (var material in materials)
                {
                    if (material.Key.IndexOf("metal", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"Found metal material: {material.Key}");
                        return material.Value;
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine($"No match found for {string.Join(", ", possibleNames)}");
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
            // First check if the element has a "Phasing" parameter and if it's set to "proposed"
            Parameter phasingParam = element.LookupParameter("Phasing");
            if (phasingParam != null && phasingParam.HasValue)
            {
                string phasing = phasingParam.AsString();
                if (phasing == null || phasing.ToLower() != "proposed")
                {
                    // If the Phasing parameter exists but is not "proposed", return null
                    return null;
                }
            }
            // If no Phasing parameter exists, continue with detection (optional - you can also return null here)

            // First check for piles - do this first to catch all pile elements
            // Check name, family name, category, type name, and comments for pile-related keywords
            string elementName = element.Name?.ToLower() ?? "";
            string familyName = "";
            string comments = element.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsString()?.ToLower() ?? "";
            string typeName = "";

            // Get family name if it's a family instance
            if (element is FamilyInstance familyInstance)
            {
                // Get the family name
                try
                {
                    ElementId typeId = element.GetTypeId();
                    if (typeId != ElementId.InvalidElementId)
                    {
                        ElementType type = element.Document.GetElement(typeId) as ElementType;
                        if (type != null)
                        {
                            typeName = type.Name?.ToLower() ?? "";

                            // Get family name if available
                            try
                            {
                                familyName = type.FamilyName?.ToLower() ?? "";
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }

            // Check for pile-related keywords in various properties
            string[] pileKeywords = new string[] { "pile", "piling", "piles", "caisson", "shaft" };
            bool isPile = pileKeywords.Any(keyword =>
                elementName.Contains(keyword) ||
                familyName.Contains(keyword) ||
                typeName.Contains(keyword) ||
                comments.Contains(keyword));

            if (isPile)
            {
                return "Pilings";
            }

            // Continue with the original logic for other structural elements
            if (element is FamilyInstance familyInst)
            {
                var structType = familyInst.StructuralType;
                if (structType == StructuralType.Beam)
                    return "Beams";
                if (structType == StructuralType.Column)
                    return "Columns";
                if (structType == StructuralType.Footing)
                    return "Foundations";
            }
            else if (element is Wall wall)
            {
                // Check if it's an upstand wall
                if (comments.Contains("upstand") || wall.Name.ToLower().Contains("upstand"))
                    return "Upstands";

                return "Walls";
            }
            else if (element is Floor)
            {
                return "Floors";
            }
            else if (element is FootPrintRoof)
            {
                return "Floors"; // Group roofs with floors for now
            }

            // For foundations
            if (element.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFoundation)
                return "Foundations";

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
            try
            {
                // SIMPLE DIRECT APPROACH:
                // 1. Get concrete carbon from _items
                double concreteCarbon = _items.Values.Sum(item => item.ConcreteCarbon);

                // 2. Get total carbon from steel table (Virgin + Reused)
                double steelTotalCarbon = 0;

                if (_flattenedSteelTable != null)
                {
                    // Get Virgin Steel total
                    var virginSteel = _flattenedSteelTable.FirstOrDefault(r => r.Level == 0 && r.Name == "Virgin Steel");
                    if (virginSteel != null)
                    {
                        steelTotalCarbon += virginSteel.Carbon;
                    }

                    // Get Reused Steel total
                    var reusedSteel = _flattenedSteelTable.FirstOrDefault(r => r.Level == 0 && r.Name == "Reused Steel");
                    if (reusedSteel != null)
                    {
                        steelTotalCarbon += reusedSteel.Carbon;
                    }
                }

                // 3. Get timber and masonry carbon from _items
                double timberCarbon = _items.Values.Sum(item => item.TimberCarbon);
                double masonryCarbon = _items.Values.Sum(item => item.MasonryCarbon);

                // 4. Calculate new total including concrete
                double newTotal = (concreteCarbon + steelTotalCarbon + timberCarbon + masonryCarbon) / 1000; // Convert from kg to tonnes

                // 4. Update property
                TotalCarbonEmbodiment = newTotal;

                // 5. Notify UI
                OnPropertyChanged(nameof(TotalCarbonEmbodiment));

                // 6. Update rating
                UpdateCarbonRating();

                // 7. Notify rating properties
                OnPropertyChanged(nameof(CarbonRating));
                OnPropertyChanged(nameof(CarbonGrade));
                OnPropertyChanged(nameof(CarbonRatingColor));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in UpdateTotalCarbonEmbodiment: {ex.Message}");
            }
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

                // Create the XML document
                XDocument doc = new XDocument(
                    new XElement("ECATPluginData",
                        new XElement("ProjectName", ProjectName),
                        new XElement("ProjectNumber", ProjectNumber),
                        new XElement("Phase", Phase),
                        new XElement("GIA", GIA),
                        new XElement("GIACalculated", GIACalculated),
                        new XElement("IsInputGIASelected", IsInputGIASelected),
                        new XElement("IsCalculatedGIASelected", IsCalculatedGIASelected),

                        // Save traditional Items data
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
                                    new XElement("TimberECFactor", item.Value.TimberECFactor),
                                    new XElement("MasonryECFactor", item.Value.MasonryECFactor)
                                )
                            )
                        ),

                       // Save Steel Table data - save all user inputs
                       new XElement("SteelTable",
                    _flattenedSteelTable != null ?
                    _flattenedSteelTable.Where(row => row.IsEditable).Select(row =>
                        new XElement("Row",
                            new XAttribute("Name", row.Name),
                            new XAttribute("Level", row.Level),
                            new XAttribute("ParentElement", GetElementTypeFromRow(row)),
                            new XElement("SectionType", row.SectionType),
                            new XElement("Volume", row.Volume),
                            new XElement("Source", row.Source),
                            new XElement("CarbonFactor", row.CarbonFactor),
                            new XElement("IsManualECFactor", row.IsManualECFactor) // Add this line

                            )

                        ) : null
                    ),

                        // Save Timber Table data
                        new XElement("TimberTable",
                        TimberItemsView != null ?
                        TimberItemsView.Select(item =>
                            new XElement("Row",
                                new XAttribute("ElementType", item.Key),
                                new XElement("Volume", item.Value.TimberVolume),
                                new XElement("Type", item.Value.TimberType),
                                new XElement("Source", item.Value.TimberSource),
                                new XElement("ECFactor", item.Value.TimberECFactor),
                                new XElement("IsManualECFactor", item.Value.IsManualTimberECFactor) // Add this line
                            )
                        ) : null
                    ),


                    // Save Masonry Table data
                    new XElement("MasonryTable",
                        _flattenedMasonryTable != null ?
                        _flattenedMasonryTable.Where(row => row.IsEditable).Select(row =>
                            new XElement("Row",
                                new XAttribute("Name", row.Name),
                                new XAttribute("Level", row.Level),
                                new XElement("Volume", row.Volume),
                                new XElement("MasonryType", row.MasonryType),
                                new XElement("CarbonFactor", row.CarbonFactor),
                                new XElement("IsManualECFactor", row.IsManualECFactor)
                            )
                        ) : null
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

                                // Load volume values
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

                                // Load custom EC factors for timber and masonry
                                if (itemElement.Element("TimberECFactor") != null)
                                    itemData.TimberECFactor = double.Parse(itemElement.Element("TimberECFactor").Value);

                                if (itemElement.Element("MasonryECFactor") != null)
                                    itemData.MasonryECFactor = double.Parse(itemElement.Element("MasonryECFactor").Value);
                            }
                        }
                    }

                    // Load Steel Table data if available
                    var steelTableElement = root.Element("SteelTable");
                    if (steelTableElement != null)
                    {
                        // Make sure the flattened table is created first
                        if (_flattenedSteelTable == null)
                            _flattenedSteelTable = CreateFlattenedSteelTable();

                        // Process each saved row
                        foreach (var rowElement in steelTableElement.Elements("Row"))
                        {
                            // Get attributes
                            string name = rowElement.Attribute("Name")?.Value;
                            int level = int.Parse(rowElement.Attribute("Level")?.Value ?? "2");
                            string parentElement = rowElement.Attribute("ParentElement")?.Value;

                            // Find matching row in flattened table
                            var matchingRows = _flattenedSteelTable.Where(r =>
                                r.IsEditable && r.Name == name && GetElementTypeFromRow(r) == parentElement).ToList();

                            if (matchingRows.Count > 0)
                            {
                                var row = matchingRows[0]; // Take the first match

                                // Load values
                                if (rowElement.Element("Volume") != null)
                                    row.Volume = double.Parse(rowElement.Element("Volume").Value);

                                if (rowElement.Element("Source") != null)
                                    row.Source = rowElement.Element("Source").Value;

                                if (rowElement.Element("SectionType") != null)
                                    row.SectionType = rowElement.Element("SectionType").Value;

                                // IMPORTANT: First load the manual flag, THEN the carbon factor
                                bool isManual = false;
                                if (rowElement.Element("IsManualECFactor") != null)
                                    isManual = bool.Parse(rowElement.Element("IsManualECFactor").Value);

                                row.IsManualECFactor = isManual;

                                // Load carbon factor, using our special method to preserve the manual flag
                                if (rowElement.Element("CarbonFactor") != null)
                                {
                                    double carbonFactorValue = double.Parse(rowElement.Element("CarbonFactor").Value);
                                    row.SetCarbonFactorDirectly(carbonFactorValue, true);
                                }
                            }
                        }

                        // Update all the totals based on the loaded data
                        SyncSteelTableChanges();
                    }

                    var timberTableElement = root.Element("TimberTable");
                    if (timberTableElement != null)
                    {
                        foreach (var rowElement in timberTableElement.Elements("Row"))
                        {
                            string elementType = rowElement.Attribute("ElementType")?.Value;

                            // Find the corresponding item in the Items dictionary
                            if (!string.IsNullOrEmpty(elementType) && _items.ContainsKey(elementType))
                            {
                                var item = _items[elementType];

                                // Load values
                                if (rowElement.Element("Volume") != null)
                                    item.TimberVolume = double.Parse(rowElement.Element("Volume").Value);

                                if (rowElement.Element("Type") != null)
                                    item.TimberType = rowElement.Element("Type").Value;

                                if (rowElement.Element("ECFactor") != null)
                                    item.TimberECFactor = double.Parse(rowElement.Element("ECFactor").Value);

                                // IMPORTANT: Also load the timber source if available
                                if (rowElement.Element("Source") != null)
                                    item.TimberSource = rowElement.Element("Source").Value;

                                if (rowElement.Element("IsManualECFactor") != null)
                                    item.IsManualTimberECFactor = bool.Parse(rowElement.Element("IsManualECFactor").Value);

                                // Notify property changes
                                item.NotifyPropertyChanged("TimberVolume");
                                item.NotifyPropertyChanged("TimberType");
                                item.NotifyPropertyChanged("TimberSource"); // Add notification for source
                                item.NotifyPropertyChanged("TimberECFactor");
                                item.NotifyPropertyChanged("TimberCarbon");
                                item.NotifyPropertyChanged("SubTotalCarbon");
                            }
                        }
                    }

                    // Load Masonry Table data
                    var masonryTableElement = root.Element("MasonryTable");
                    if (masonryTableElement != null)
                    {
                        // Make sure the flattened table is created first
                        if (_flattenedMasonryTable == null)
                            _flattenedMasonryTable = CreateFlattenedMasonryTable();

                        // Process each saved row
                        foreach (var rowElement in masonryTableElement.Elements("Row"))
                        {
                            // Get attributes
                            string name = rowElement.Attribute("Name")?.Value;
                            int level = int.Parse(rowElement.Attribute("Level")?.Value ?? "2");

                            // Find matching row in flattened table
                            var matchingRows = _flattenedMasonryTable.Where(r =>
                                r.IsEditable && r.Name == name).ToList();

                            if (matchingRows.Count > 0)
                            {
                                var row = matchingRows[0]; // Take the first match

                                // Load values
                                if (rowElement.Element("Volume") != null)
                                    row.Volume = double.Parse(rowElement.Element("Volume").Value);

                                if (rowElement.Element("MasonryType") != null)
                                    row.MasonryType = rowElement.Element("MasonryType").Value;

                                // IMPORTANT: First load the manual flag, THEN the carbon factor
                                bool isManual = false;
                                if (rowElement.Element("IsManualECFactor") != null)
                                    isManual = bool.Parse(rowElement.Element("IsManualECFactor").Value);

                                row.IsManualECFactor = isManual;

                                // Load carbon factor, using our special method to preserve the manual flag
                                if (rowElement.Element("CarbonFactor") != null)
                                {
                                    double carbonFactorValue = double.Parse(rowElement.Element("CarbonFactor").Value);
                                    row.SetCarbonFactorDirectly(carbonFactorValue, true);
                                }
                            }
                        }

                        // Update all the totals based on the loaded data
                        SyncMasonryTableChanges();
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
                            item.Key == "Beams" ||
                            item.Key == "Columns" ||
                            item.Key == "Floors" ||
                            item.Key == "Pilings")
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
                            item.Key == "Beams" ||
                            item.Key == "Columns" ||
                            item.Key == "Walls" ||
                            item.Key == "Floors")
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
                        _items.Where(item => item.Key == "Walls")
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
                    return itemKey == "Beams" ||
                           itemKey == "Columns" ||
                           itemKey == "Floors" ||
                           itemKey == "Pilings";
                case "Timber":
                    return itemKey == "Beams" ||
                           itemKey == "Columns" ||
                           itemKey == "Walls" ||
                           itemKey == "Floors";
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

        // Update the RefreshVolumes method in MainViewModel.cs to ensure it properly 
        // reloads all data from the Revit model

        public void RefreshVolumes(UIApplication uiApp = null)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Starting volume refresh from Revit model...");

                // Reset all volumes to zero first to ensure clean data
                foreach (var key in _items.Keys)
                {
                    _items[key].ConcreteVolume = 0;
                    _items[key].SteelVolume = 0;
                    _items[key].TimberVolume = 0;
                    _items[key].MasonryVolume = 0;
                }

                // Reload the document reference to ensure we're working with the current state
                Document currentDoc = _document;
                if (uiApp != null && uiApp.ActiveUIDocument != null)
                {
                    currentDoc = uiApp.ActiveUIDocument.Document;
                }

                // Reload GIA in case floor areas have changed
                LoadGIA();

                // Then reload volumes from the current document state - with detailed reporting
                System.Diagnostics.Debug.WriteLine("Loading materials from Revit model...");

                // Get all materials in the document first
                var allMaterials = new FilteredElementCollector(currentDoc)
                    .OfClass(typeof(Material))
                    .Cast<Material>()
                    .ToDictionary(m => m.Name, m => m.Id);

                // Report found materials
                System.Diagnostics.Debug.WriteLine($"Found {allMaterials.Count} materials in the model:");
                foreach (var mat in allMaterials.Take(10)) // Log first 10 for debugging
                {
                    System.Diagnostics.Debug.WriteLine($"  - {mat.Key}");
                }

                // Find the closest matching materials
                ElementId concreteMatId = FindBestMaterialMatch(allMaterials, "Concrete", "WAL_Concrete_RC");
                ElementId steelMatId = FindBestMaterialMatch(allMaterials, "Steel", "WAL_Steel", "Metal");
                ElementId timberMatId = FindBestMaterialMatch(allMaterials, "Wood", "Wood - Dimensional Lumber", "Lumber");
                ElementId blockMatId = FindBestMaterialMatch(allMaterials, "Block", "WAL_Block", "Masonry");
                ElementId brickMatId = FindBestMaterialMatch(allMaterials, "Brick", "WAL_Brick");

                System.Diagnostics.Debug.WriteLine("Material mapping results:");
                System.Diagnostics.Debug.WriteLine($"  - Concrete material: {(concreteMatId != null ? "Found" : "Not found")}");
                System.Diagnostics.Debug.WriteLine($"  - Steel material: {(steelMatId != null ? "Found" : "Not found")}");
                System.Diagnostics.Debug.WriteLine($"  - Timber material: {(timberMatId != null ? "Found" : "Not found")}");
                System.Diagnostics.Debug.WriteLine($"  - Block material: {(blockMatId != null ? "Found" : "Not found")}");
                System.Diagnostics.Debug.WriteLine($"  - Brick material: {(brickMatId != null ? "Found" : "Not found")}");

                // Initialize the steel volume by section dictionary
                _steelVolumeBySection = new Dictionary<string, Dictionary<string, double>>();
                foreach (var key in _items.Keys)
                {
                    if (IsItemValidForMaterial(key, "Steel"))
                    {
                        _steelVolumeBySection[key] = new Dictionary<string, double>
                {
                    { "Open Sections", 0 },
                    { "Closed Sections", 0 },
                    { "Plates", 0 }
                };
                    }
                }

                // Get all structural elements
                var structuralElements = GetAllStructuralElements(currentDoc);
                System.Diagnostics.Debug.WriteLine($"Found {structuralElements.Count} structural elements");

                // Dictionary to store volumes by element type
                var concreteVolumes = new Dictionary<string, double>();
                var steelVolumes = new Dictionary<string, double>();
                var timberVolumes = new Dictionary<string, double>();
                var brickVolumes = new Dictionary<string, double>();
                var blockVolumes = new Dictionary<string, double>();

                // Track processed elements for reporting
                int processedCount = 0;
                int skippedCount = 0;
                int concreteCount = 0;
                int steelCount = 0;
                int timberCount = 0;
                int brickCount = 0;
                int blockCount = 0;

                // Process each element
                foreach (var element in structuralElements)
                {
                    try
                    {
                        // Check for Phasing parameter - only process elements in the "Proposed" phase
                        Parameter phasingParam = element.LookupParameter("Phasing");
                        if (phasingParam != null && phasingParam.HasValue)
                        {
                            string phasing = phasingParam.AsString();
                            if (phasing == null || !phasing.Equals("Proposed", StringComparison.OrdinalIgnoreCase))
                            {
                                skippedCount++;
                                continue; // Skip if not "Proposed"
                            }
                        }

                        // Skip elements without materials
                        var materials = element.GetMaterialIds(false);
                        if (materials == null || materials.Count == 0)
                        {
                            skippedCount++;
                            continue;
                        }

                        string elementType = GetElementType(element);
                        if (string.IsNullOrEmpty(elementType))
                        {
                            skippedCount++;
                            continue;
                        }

                        bool elementProcessed = false;

                        // Process each material in the element
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

                            elementProcessed = true;
                            processedCount++;

                            // Concrete is applicable for all structural elements
                            if (materialId.Equals(concreteMatId))
                            {
                                if (!concreteVolumes.ContainsKey(elementType))
                                    concreteVolumes[elementType] = 0;
                                concreteVolumes[elementType] += volume;
                                concreteCount++;
                            }
                            // Steel is only for beams, columns, floors, foundations, and pilings
                            else if (materialId.Equals(steelMatId) && IsItemValidForMaterial(elementType, "Steel"))
                            {
                                if (!steelVolumes.ContainsKey(elementType))
                                    steelVolumes[elementType] = 0;
                                steelVolumes[elementType] += volume;
                                steelCount++;

                                // Determine steel section type ONLY for elements with steel material ID
                                if (_steelVolumeBySection.ContainsKey(elementType))
                                {
                                    // Get the section type for this steel element
                                    string sectionType = DetermineSectionType(element);

                                    // Only add volume if the section type was identified
                                    if (!string.IsNullOrEmpty(sectionType))
                                    {
                                        // Add the volume to the appropriate section category
                                        _steelVolumeBySection[elementType][sectionType] += volume;
                                    }
                                }
                            }
                            // Timber is only for beams, columns, walls, and floors
                            else if (materialId.Equals(timberMatId) && IsItemValidForMaterial(elementType, "Timber"))
                            {
                                if (!timberVolumes.ContainsKey(elementType))
                                    timberVolumes[elementType] = 0;
                                timberVolumes[elementType] += volume;
                                timberCount++;
                            }
                            // Brick is only for walls
                            else if (materialId.Equals(brickMatId) && elementType == "Walls")
                            {
                                if (!brickVolumes.ContainsKey(elementType))
                                    brickVolumes[elementType] = 0;
                                brickVolumes[elementType] += volume;
                                brickCount++;
                            }
                            // Block is only for walls
                            else if (materialId.Equals(blockMatId) && elementType == "Walls")
                            {
                                if (!blockVolumes.ContainsKey(elementType))
                                    blockVolumes[elementType] = 0;
                                blockVolumes[elementType] += volume;
                                blockCount++;
                            }
                        }

                        if (!elementProcessed)
                        {
                            skippedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error processing element {element.Id}: {ex.Message}");
                        skippedCount++;
                    }
                }

                // Log processing statistics
                System.Diagnostics.Debug.WriteLine($"Elements processing summary:");
                System.Diagnostics.Debug.WriteLine($"  - Processed: {processedCount}");
                System.Diagnostics.Debug.WriteLine($"  - Skipped: {skippedCount}");
                System.Diagnostics.Debug.WriteLine($"  - With concrete: {concreteCount}");
                System.Diagnostics.Debug.WriteLine($"  - With steel: {steelCount}");
                System.Diagnostics.Debug.WriteLine($"  - With timber: {timberCount}");
                System.Diagnostics.Debug.WriteLine($"  - With brick: {brickCount}");
                System.Diagnostics.Debug.WriteLine($"  - With block: {blockCount}");

                // Update the dictionary values with cubic meter conversion
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

                    // Total masonry volume is combined brick and block for walls
                    if (key == "Walls")
                    {
                        double brickVol = brickVolumes.ContainsKey(key) ? brickVolumes[key] / 35.315 : 0;
                        double blockVol = blockVolumes.ContainsKey(key) ? blockVolumes[key] / 35.315 : 0;
                        _items[key].MasonryVolume = brickVol + blockVol;

                        // Store the individual volumes for FlattenedMasonryTable if it exists
                        if (_flattenedMasonryTable != null)
                        {
                            var brickRow = _flattenedMasonryTable.FirstOrDefault(r => r.IsEditable && r.Name == "Brickwork");
                            if (brickRow != null)
                                brickRow.Volume = brickVol;

                            var blockRow = _flattenedMasonryTable.FirstOrDefault(r => r.IsEditable && r.Name == "Blockwork");
                            if (blockRow != null)
                                blockRow.Volume = blockVol;

                            // Sync masonry table changes to update totals
                            SyncMasonryTableChanges();
                        }
                    }
                    else
                    {
                        _items[key].MasonryVolume = 0;
                    }

                    System.Diagnostics.Debug.WriteLine($"Updated {key}: Concrete={_items[key].ConcreteVolume:F2}, Steel={_items[key].SteelVolume:F2}, Timber={_items[key].TimberVolume:F2}, Masonry={_items[key].MasonryVolume:F2}");
                }

                // Update carbon calculations
                UpdateTotalCarbonEmbodiment();

                // Update UI collections
                UpdateItemsView();

                // Reset and rebuild hierarchical tables
                _steelHierarchy = null;
                OnPropertyChanged(nameof(SteelHierarchy));

                _flattenedSteelTable = null;
                OnPropertyChanged(nameof(FlattenedSteelTable));

                // Only reset masonry table if it exists
                if (_flattenedMasonryTable != null)
                {
                    _flattenedMasonryTable = null;
                    OnPropertyChanged(nameof(FlattenedMasonryTable));
                }

                System.Diagnostics.Debug.WriteLine("Volume refresh completed successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error during RefreshVolumes: {ex.Message}");
                throw; // Rethrow to show error message to user
            }
        }

        // Steel calculations



        // Improved method to determine the section type from a family name


        // Create a new class for representing steel section categories
        public class SteelSectionCategory : INotifyPropertyChanged
        {
            public string Name { get; set; } // e.g., "Open Sections", "Closed Sections", "Plates"

            private double _volume;
            public double Volume
            {
                get => _volume;
                set
                {
                    _volume = value;
                    OnPropertyChanged(nameof(Volume));
                    OnPropertyChanged(nameof(Carbon));
                }
            }

            private string _source = "UK";
            public string Source
            {
                get => _source;
                set
                {
                    _source = value;
                    UpdateCarbonFactor();
                    OnPropertyChanged(nameof(Source));
                    OnPropertyChanged(nameof(Carbon));
                }
            }

            private double _carbonFactor;
            public double CarbonFactor
            {
                get => _carbonFactor;
                set
                {
                    _carbonFactor = value;
                    OnPropertyChanged(nameof(CarbonFactor));
                    OnPropertyChanged(nameof(Carbon));
                }
            }

            private string _module = "A1-A3";
            public string Module
            {
                get => _module;
                set
                {
                    _module = value;
                    OnPropertyChanged(nameof(Module));
                }
            }
            private bool _isManualECFactor = false;
            public bool IsManualECFactor
            {
                get => _isManualECFactor;
                set
                {
                    _isManualECFactor = value;
                    OnPropertyChanged(nameof(IsManualECFactor));
                }
            }

            private string _sectionType;
            public string SectionType
            {
                get => _sectionType;
                set
                {
                    if (_sectionType != value)
                    {
                        _sectionType = value;
                        OnPropertyChanged(nameof(SectionType));

                        // Only update carbon factor if not manually set
                        if (!IsManualECFactor)
                        {
                            UpdateCarbonFactor();
                        }
                    }
                }
            }


            private Dictionary<string, Dictionary<string, double>> _steelVolumeBySection;



            // Reference to the steel data for calculating carbon factors
            private Dictionary<string, Dictionary<string, double>> _steelData;

            public SteelSectionCategory(Dictionary<string, Dictionary<string, double>> steelData)
            {
                _steelData = steelData;
            }

            private void UpdateCarbonFactor()
            {
                // Map section category to steel data keys
                string steelDataKey = _sectionType == "Open Sections" ? "Open Section" :
                                     _sectionType == "Closed Sections" ? "Closed Section" : "Plates";

                // Look up the carbon factor for this section type and source
                if (_steelData != null && _steelData.ContainsKey(steelDataKey) &&
                    _steelData[steelDataKey].ContainsKey(_source))
                {
                    CarbonFactor = _steelData[steelDataKey][_source];
                }
            }

            // Calculate carbon based on volume and factor
            public double Carbon => Volume * CarbonFactor * 7850; // Steel density = 7850 kg/m³

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // Create a new class for representing steel elements
        public class SteelElementGroup : INotifyPropertyChanged
        {
            public string Name { get; set; } // e.g., "Beams", "Columns", "Floors", "Piles"

            private ObservableCollection<SteelSectionCategory> _sections = new ObservableCollection<SteelSectionCategory>();
            public ObservableCollection<SteelSectionCategory> Sections
            {
                get => _sections;
                set
                {
                    _sections = value;
                    OnPropertyChanged(nameof(Sections));
                    OnPropertyChanged(nameof(TotalCarbon));
                    OnPropertyChanged(nameof(TotalVolume));
                }
            }

            // Total carbon for this element group
            public double TotalCarbon => Sections.Sum(s => s.Carbon);

            // Total volume for this element group
            public double TotalVolume => Sections.Sum(s => s.Volume);

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // Create a class to represent steel by source type
        public class SteelSourceGroup : INotifyPropertyChanged
        {
            public string Name { get; set; } // e.g., "Virgin Steel", "Reused Steel"

            private ObservableCollection<SteelElementGroup> _elements = new ObservableCollection<SteelElementGroup>();
            public ObservableCollection<SteelElementGroup> Elements
            {
                get => _elements;
                set
                {
                    _elements = value;
                    OnPropertyChanged(nameof(Elements));
                    OnPropertyChanged(nameof(TotalCarbon));
                    OnPropertyChanged(nameof(TotalVolume));
                }
            }

            // Total carbon for this source group
            public double TotalCarbon => Elements.Sum(e => e.TotalCarbon);

            // Total volume for this source group
            public double TotalVolume => Elements.Sum(e => e.TotalVolume);

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // Add a new method to MainViewModel to collect volumes by section type
        // This will be called during LoadVolumes
        // Enhanced method to collect steel volumes by section for WAL_SFRA_STL components
        private Dictionary<string, Dictionary<string, double>> CollectSteelVolumesBySection(Document doc)
        {
            // Structure: Element Type => Section Type => Volume
            var volumes = new Dictionary<string, Dictionary<string, double>>();

            // Initialize for all element types
            foreach (var key in _items.Keys)
            {
                volumes[key] = new Dictionary<string, double>
        {
            { "Open Sections", 0 },
            { "Closed Sections", 0 },
            { "Plates", 0 }
        };
            }

            System.Diagnostics.Debug.WriteLine("Starting steel volume collection...");

            // Get structural framing elements (beams)
            var structuralFraming = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .WhereElementIsNotElementType()
                .ToElements();

            System.Diagnostics.Debug.WriteLine($"Found {structuralFraming.Count()} structural framing elements");

            // Get structural column elements
            var structuralColumns = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .WhereElementIsNotElementType()
                .ToElements();

            System.Diagnostics.Debug.WriteLine($"Found {structuralColumns.Count()} structural column elements");

            // Process all structural elements
            var allElements = new List<Element>();
            allElements.AddRange(structuralFraming);
            allElements.AddRange(structuralColumns);

            foreach (var element in allElements)
            {
                try
                {
                    // Check for Phasing parameter - skip if not "Proposed"
                    Parameter phasingParam = element.LookupParameter("Phasing");
                    if (phasingParam != null && phasingParam.HasValue)
                    {
                        string phasing = phasingParam.AsString();
                        if (phasing == null || !phasing.Equals("Proposed", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }
                    }

                    // Determine element type (Beams or Columns)
                    string elementType = "";
                    if (element is FamilyInstance familyInstance)
                    {
                        var structType = familyInstance.StructuralType;
                        if (structType == StructuralType.Beam)
                            elementType = "Beams";
                        else if (structType == StructuralType.Column)
                            elementType = "Columns";
                        else
                            continue; // Skip if not a beam or column
                    }
                    else
                    {
                        continue; // Skip if not a family instance
                    }

                    if (string.IsNullOrEmpty(elementType) || !volumes.ContainsKey(elementType))
                    {
                        continue;
                    }

                    // Calculate volume using the existing CalculateElementVolume method
                    double volume = 0;
                    try
                    {
                        volume = CalculateElementVolume(element);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error calculating volume for element: {ex.Message}");
                        continue;
                    }

                    if (volume <= 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"Zero or negative volume for element");
                        continue;
                    }

                    // Determine section type using the DetermineSectionType method
                    string sectionType = DetermineSectionType(element);

                    // Add the volume to the appropriate section type
                    volumes[elementType][sectionType] += volume;

                    System.Diagnostics.Debug.WriteLine($"Added {volume:F3} cubic feet of {sectionType} to {elementType}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error processing element: {ex.Message}");
                }
            }

            // Convert cubic feet to cubic meters
            foreach (var elemType in volumes.Keys)
            {
                foreach (var sectionType in volumes[elemType].Keys.ToList())
                {
                    volumes[elemType][sectionType] /= 35.315;
                }
            }

            return volumes;
        }



        // Helper method to determine element type specifically for steel elements
        private string GetSteelElementType(Element element)
        {
            if (element is FamilyInstance familyInstance)
            {
                var structType = familyInstance.StructuralType;

                if (structType == StructuralType.Beam)
                    return "Beams";
                if (structType == StructuralType.Column)
                    return "Columns";
            }

            // Try to determine by category as a fallback
            if (element.Category?.Name != null)
            {
                string categoryName = element.Category.Name.ToLower();
                if (categoryName.Contains("framing") || categoryName.Contains("beam"))
                    return "Beams";
                if (categoryName.Contains("column"))
                    return "Columns";
            }

            return null;
        }

        // Add a property to MainViewModel for the hierarchical steel data
        private ObservableCollection<SteelSourceGroup> _steelHierarchy;
        public ObservableCollection<SteelSourceGroup> SteelHierarchy
        {
            get
            {
                if (_steelHierarchy == null)
                {
                    _steelHierarchy = BuildSteelHierarchy();
                }
                else if (_steelVolumeBySection != null)
                {
                    // Update existing hierarchy with new volumes
                    foreach (var sourceGroup in _steelHierarchy)
                    {
                        foreach (var elementGroup in sourceGroup.Elements)
                        {
                            string elementType = elementGroup.Name;
                            if (_steelVolumeBySection.ContainsKey(elementType))
                            {
                                // Update sections with the categorized volumes
                                foreach (var section in elementGroup.Sections)
                                {
                                    if (_steelVolumeBySection[elementType].ContainsKey(section.Name))
                                    {
                                        // Convert from cubic feet to cubic meters
                                        section.Volume = _steelVolumeBySection[elementType][section.Name] / 35.315;
                                    }
                                }
                            }
                        }
                    }
                }
                return _steelHierarchy;
            }
        }

        // Method to build the hierarchical steel data structure
        private ObservableCollection<SteelSourceGroup> BuildSteelHierarchy()
        {
            var hierarchy = new ObservableCollection<SteelSourceGroup>();

            // Create Virgin Steel group
            var virginSteel = new SteelSourceGroup { Name = "Virgin Steel" };
            virginSteel.Elements = new ObservableCollection<SteelElementGroup>();

            // Create Reused Steel group
            var reusedSteel = new SteelSourceGroup { Name = "Reused Steel" };
            reusedSteel.Elements = new ObservableCollection<SteelElementGroup>();

            // Create element groups and add to both Virgin and Reused steel
            foreach (var elementType in new[] { "Beams", "Columns", "Floors", "Pilings" })
            {
                // Skip element types not in our data
                if (_steelVolumeBySection == null || !_steelVolumeBySection.ContainsKey(elementType))
                    continue;

                // Create element group for Virgin Steel
                var virginElement = new SteelElementGroup { Name = elementType };
                virginElement.Sections = new ObservableCollection<SteelSectionCategory>();

                // Create element group for Reused Steel
                var reusedElement = new SteelElementGroup { Name = elementType };
                reusedElement.Sections = new ObservableCollection<SteelSectionCategory>();

                // Only add section categories for Beams and Columns
                if (elementType == "Beams" || elementType == "Columns")
                {
                    // Add section categories for Open Sections, Closed Sections, and Plates
                    foreach (var sectionType in new[] { "Open Sections", "Closed Sections", "Plates" })
                    {
                        // Get volume for this section type (convert from cubic feet to cubic meters)
                        double volume = _steelVolumeBySection[elementType][sectionType] / 35.315;

                        // Add to Virgin Steel
                        var virginSection = new SteelSectionCategory(_steelData)
                        {
                            Name = sectionType,
                            SectionType = sectionType,
                            Volume = volume,
                            Source = "UK", // Default source
                            Module = "A1-A3" // Default module
                        };
                        virginElement.Sections.Add(virginSection);

                        // Add to Reused Steel (with zero volume for now)
                        var reusedSection = new SteelSectionCategory(_steelData)
                        {
                            Name = sectionType,
                            SectionType = sectionType,
                            Volume = 0, // No reused volume by default
                            Source = "UK (Reused)",
                            Module = "A1-A3"
                        };
                        reusedElement.Sections.Add(reusedSection);
                    }
                }
                else
                {
                    // For Floors and Piles, add a single entry without section subcategories
                    double totalVolume = _steelVolumeBySection[elementType].Values.Sum() / 35.315;

                    // Add to Virgin Steel
                    virginElement.Sections.Add(new SteelSectionCategory(_steelData)
                    {
                        Name = elementType,
                        SectionType = "Open Sections", // Default section type
                        Volume = totalVolume,
                        Source = "UK",
                        Module = "A1-A3"
                    });

                    // Add to Reused Steel (with zero volume)
                    reusedElement.Sections.Add(new SteelSectionCategory(_steelData)
                    {
                        Name = elementType,
                        SectionType = "Open Sections", // Default section type
                        Volume = 0,
                        Source = "UK (Reused)",
                        Module = "A1-A3"
                    });
                }

                // Add element groups to source groups if they have volume
                if (virginElement.TotalVolume > 0)
                    virginSteel.Elements.Add(virginElement);

                if (reusedElement.TotalVolume > 0)
                    reusedSteel.Elements.Add(reusedElement);
            }

            // Add source groups to hierarchy
            hierarchy.Add(virginSteel);
            hierarchy.Add(reusedSteel);

            return hierarchy;
        }

        // Changes needed in MainViewModel.cs

        // Update the SteelTableRow class by removing Module property and related code
        public class SteelTableRow : INotifyPropertyChanged
        {
            // Basic properties
            private string _name;
            public string Name
            {
                get => _name;
                set
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }

            public void SetCarbonFactorDirectly(double value, bool preserveManualFlag)
            {
                // Store the current manual flag state
                bool currentManualState = IsManualECFactor;

                // Set the carbon factor value directly
                _carbonFactor = value;

                // If we need to preserve the manual flag, restore it
                if (preserveManualFlag)
                {
                    IsManualECFactor = currentManualState;
                }

                // Notify that the properties have changed
                OnPropertyChanged(nameof(CarbonFactor));
                OnPropertyChanged(nameof(Carbon));
            }

            private int _level;
            public int Level
            {
                get => _level;
                set
                {
                    _level = value;
                    OnPropertyChanged(nameof(Level));
                    OnPropertyChanged(nameof(IndentMargin));
                    OnPropertyChanged(nameof(FontWeight));
                }
            }

            private double _volume;
            public double Volume
            {
                get => _volume;
                set
                {
                    _volume = value;
                    OnPropertyChanged(nameof(Volume));
                    OnPropertyChanged(nameof(Carbon));
                }
            }

            public void SetCarbonFactorFromDropdown(double value)
            {
                // When setting from dropdown, update the value but don't reset IsManualECFactor
                // This allows manually-entered values to remain manual even after dropdown selections
                _carbonFactor = value;
                OnPropertyChanged(nameof(CarbonFactor));
                OnPropertyChanged(nameof(Carbon));
            }

            private double _carbonFactor;
            public double CarbonFactor
            {
                get => _carbonFactor;
                set
                {
                    if (Math.Abs(_carbonFactor - value) > 0.0001)
                    {
                        double oldValue = _carbonFactor;
                        _carbonFactor = value;
                        System.Diagnostics.Debug.WriteLine($"CarbonFactor changed for {Name} from {oldValue} to {value}");

                        // Flag to indicate this was manually set
                        IsManualECFactor = true;

                        // Notify dependent properties
                        OnPropertyChanged(nameof(CarbonFactor));
                        OnPropertyChanged(nameof(Carbon));
                    }
                }
            }

            private bool _isManualECFactor = false;
            public bool IsManualECFactor
            {
                get => _isManualECFactor;
                set
                {
                    _isManualECFactor = value;
                    OnPropertyChanged(nameof(IsManualECFactor));
                }
            }

            // Removed Module property

            private bool _isEditable = false;
            public bool IsEditable
            {
                get => _isEditable;
                set
                {
                    _isEditable = value;
                    OnPropertyChanged(nameof(IsEditable));
                    OnPropertyChanged(nameof(EditableVisibility));
                    OnPropertyChanged(nameof(VolumeVisibility));
                }
            }

            private string _sectionType;
            public string SectionType
            {
                get => _sectionType;
                set
                {
                    if (_sectionType != value)
                    {
                        _sectionType = value;
                        OnPropertyChanged(nameof(SectionType));
                        UpdateCarbonFactor(); // This will update the carbon factor when section type changes
                    }
                }
            }

            private string _source = "UK";
            public string Source
            {
                get => _source;
                set
                {
                    if (_source != value)
                    {
                        _source = value;
                        OnPropertyChanged(nameof(Source));

                        // Always update Carbon calculation
                        OnPropertyChanged(nameof(Carbon));
                    }
                }
            }

            private bool _carbonFactorEditable = true;
            public bool CarbonFactorEditable
            {
                get => _carbonFactorEditable;
                set
                {
                    _carbonFactorEditable = value;
                    OnPropertyChanged(nameof(CarbonFactorEditable));
                }
            }

            // Reference to the steel data for calculating carbon factors
            private Dictionary<string, Dictionary<string, double>> _steelData;

            // Reference to main view model - can be set when row is created
            private MainViewModel _mainViewModel;

            // Helper method to get main view model - use Application.Current.MainWindow.DataContext as fallback
            private MainViewModel GetMainViewModel()
            {
                if (_mainViewModel != null)
                    return _mainViewModel;

                // Fallback to Application.Current.MainWindow.DataContext
                if (System.Windows.Application.Current?.MainWindow?.DataContext is MainViewModel viewModel)
                    return viewModel;

                return null;
            }

            // Constructor with option to store reference to view model
            public SteelTableRow(Dictionary<string, Dictionary<string, double>> steelData = null, MainViewModel mainViewModel = null)
            {
                _steelData = steelData;
                _mainViewModel = mainViewModel;
            }

            // Calculated properties
            public Thickness IndentMargin => new Thickness(Level * 20, 0, 0, 0);

            public System.Windows.FontWeight FontWeight
            {
                get
                {
                    if (Level == 0) return System.Windows.FontWeights.Bold;
                    if (Level == 1) return System.Windows.FontWeights.SemiBold;
                    return System.Windows.FontWeights.Normal;
                }
            }

            public double Carbon
            {
                get
                {
                    if (IsEditable)
                        return Volume * CarbonFactor * 7850; // Steel density = 7850 kg/m³

                    // For parent rows (non-editable), the Carbon is calculated in the view model
                    return _carbon;
                }
            }

            private double _carbon;
            public void SetCarbon(double value)
            {
                _carbon = value;
                OnPropertyChanged(nameof(Carbon));
            }

            // Visibility properties for UI binding
            public System.Windows.Visibility EditableVisibility => IsEditable ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            public System.Windows.Visibility VolumeVisibility => IsEditable ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;

            // Update the carbon factor based on section type and source
            private void UpdateCarbonFactor()
            {
                // Keep the IsManualECFactor check - this is used for programmatic updates
                if (IsManualECFactor) return;

                if (_steelData == null || string.IsNullOrEmpty(_sectionType)) return;

                // Map section category to steel data keys
                string steelDataKey = _sectionType == "Open Sections" ? "Open Section" :
                                      _sectionType == "Closed Sections" ? "Closed Section" : "Plates";

                // Look up the carbon factor for this section type and source
                if (_steelData.ContainsKey(steelDataKey) && _steelData[steelDataKey].ContainsKey(_source))
                {
                    // Only update if not manually edited
                    CarbonFactor = _steelData[steelDataKey][_source];
                    IsManualECFactor = false; // Reset the flag since we're updating from defaults
                }
            }

            // Method to recalculate carbon
            public void RecalculateCarbon()
            {
                // This will force the Carbon property to be recalculated
                OnPropertyChanged(nameof(Carbon));

                // If this is a parent row, make sure SetCarbon is called with updated value
                if (!IsEditable)
                {
                    // Get the main view model
                    var mainViewModel = GetMainViewModel();
                    if (mainViewModel != null)
                    {
                        // Get the flattened table from the main view model
                        var allRows = mainViewModel.FlattenedSteelTable;
                        if (allRows != null)
                        {
                            // Use the main view model's method to determine element type
                            var elementType = MainViewModel.GetElementTypeFromRow(this, allRows);
                            if (!string.IsNullOrEmpty(elementType))
                            {
                                // Get all child rows for this element
                                var childRows = allRows.Where(r => r.IsEditable &&
                                    MainViewModel.GetElementTypeFromRow(r, allRows) == elementType).ToList();

                                if (childRows.Any())
                                {
                                    // Sum all child carbon values
                                    double totalCarbon = childRows.Sum(r => r.Carbon);
                                    SetCarbon(totalCarbon);
                                }
                            }
                        }
                    }
                }
            }

            // INotifyPropertyChanged implementation
            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // Simplified version that just calls UpdateTotalCarbonEmbodiment directly
        public void ForceUpdateSteelCarbon()
        {
            // Simple direct approach just like the concrete tab
            UpdateTotalCarbonEmbodiment();
        }

        private void UpdateSteelHierarchyTotals()
        {
            // Skip if the flattened table hasn't been created yet
            if (_flattenedSteelTable == null)
                return;

            System.Diagnostics.Debug.WriteLine("Updating steel hierarchy totals");

            // Update element-level totals
            var elementRows = _flattenedSteelTable.Where(r => r.Level == 1).ToList();
            foreach (var elementRow in elementRows)
            {
                // Find all editable rows for this element
                var childRows = _flattenedSteelTable
                    .Where(r => r.IsEditable && GetElementTypeFromRow(r) == elementRow.Name)
                    .ToList();

                // Calculate total volume and carbon from editable rows
                double totalVolume = childRows.Sum(r => r.Volume);
                double totalCarbon = childRows.Sum(r => r.Carbon);

                // Update the element row
                elementRow.Volume = totalVolume;
                elementRow.SetCarbon(totalCarbon);

                System.Diagnostics.Debug.WriteLine($"Updated {elementRow.Name} totals: volume={totalVolume}, carbon={totalCarbon}");
            }

            // Update Virgin Steel (top-level) totals
            double virginTotalVolume = elementRows.Sum(r => r.Volume);
            double virginTotalCarbon = elementRows.Sum(r => r.Carbon);

            var virginSteel = _flattenedSteelTable.FirstOrDefault(r => r.Level == 0 && r.Name == "Virgin Steel");
            if (virginSteel != null)
            {
                virginSteel.Volume = virginTotalVolume;
                virginSteel.SetCarbon(virginTotalCarbon);
                System.Diagnostics.Debug.WriteLine($"Updated Virgin Steel totals: volume={virginTotalVolume}, carbon={virginTotalCarbon}");
            }
        }

        // Sync full steel model from UI to data model
        private void SyncFullSteelModel()
        {
            // Skip if the flattened table hasn't been created yet
            if (_flattenedSteelTable == null)
                return;

            System.Diagnostics.Debug.WriteLine("Syncing full steel model from UI to data model");

            // Go through each element type that should have steel
            foreach (var elementType in new[] { "Beams", "Columns", "Floors", "Pilings" })
            {
                if (!_items.ContainsKey(elementType))
                    continue;

                // Find the element row in the flattened table
                var elementRow = _flattenedSteelTable.FirstOrDefault(r => r.Level == 1 && r.Name == elementType);
                if (elementRow == null)
                    continue;

                // Get all child rows for this element
                var childRows = _flattenedSteelTable
                    .Where(r => r.IsEditable && GetElementTypeFromRow(r) == elementType)
                    .ToList();

                if (!childRows.Any())
                    continue;

                // Update the volume in the data model
                _items[elementType].SteelVolume = elementRow.Volume;

                // Find the dominant row (with highest volume) to set type and source
                var dominantRow = childRows.OrderByDescending(r => r.Volume).First();

                // Update steel properties in the data model
                _items[elementType].SteelType = ConvertSectionTypeToSteelType(dominantRow.SectionType);
                _items[elementType].SteelSource = dominantRow.Source;

                System.Diagnostics.Debug.WriteLine($"Updated {elementType} in data model: volume={elementRow.Volume}, type={_items[elementType].SteelType}, source={_items[elementType].SteelSource}");
            }

            // Force a property change notification for all items
            foreach (var item in _items.Values)
            {
                if (item is INotifyPropertyChanged notifyItem)
                {
                    // Try to force property change for SteelCarbon
                    var method = notifyItem.GetType().GetMethod("OnPropertyChanged",
                        System.Reflection.BindingFlags.Instance |
                        System.Reflection.BindingFlags.Public |
                        System.Reflection.BindingFlags.NonPublic);

                    if (method != null)
                    {
                        method.Invoke(notifyItem, new object[] { "SteelCarbon" });
                        method.Invoke(notifyItem, new object[] { "SubTotalCarbon" });
                    }
                }
            }
            foreach (var item in _items.Values)
            {
                item.NotifyPropertyChanged(nameof(ItemData.SteelCarbon));
                item.NotifyPropertyChanged(nameof(ItemData.SubTotalCarbon));
            }
        }

        // Add the GetElementTypeFromRow static method
        public static string GetElementTypeFromRow(SteelTableRow row, ObservableCollection<SteelTableRow> allRows = null)
        {
            if (row.Level == 1)
            {
                // This is already an element row
                return row.Name;
            }

            // For section rows, we need to find the parent element row
            if (row.Level == 2 && allRows != null)
            {
                // Find the parent row (level 1) that this section belongs to
                int rowIndex = allRows.IndexOf(row);
                if (rowIndex <= 0) return null;

                // Look backward to find the parent row
                for (int i = rowIndex - 1; i >= 0; i--)
                {
                    if (allRows[i].Level == 1)
                    {
                        // Found the parent row
                        return allRows[i].Name;
                    }
                }
            }

            return null;
        }

        // Add this method to MainViewModel to create a flattened table structure
        private ObservableCollection<SteelTableRow> _flattenedSteelTable;
        public ObservableCollection<SteelTableRow> FlattenedSteelTable
        {
            get
            {
                if (_flattenedSteelTable == null)
                {
                    _flattenedSteelTable = CreateFlattenedSteelTable();
                }
                return _flattenedSteelTable;
            }
        }

        // Method to create a flattened table from the hierarchical steel data
        // Update the CreateFlattenedSteelTable method to remove Module
        private ObservableCollection<SteelTableRow> CreateFlattenedSteelTable()
        {
            var result = new ObservableCollection<SteelTableRow>();
            var sectionVolumes = CollectSteelVolumesBySection(_document);

            // Create Virgin Steel parent row
            var virginSteel = new SteelTableRow(_steelData, this)
            {
                Name = "Virgin Steel",
                Level = 0,
                IsEditable = false,
                IsManualECFactor = false
            };

            // Create Reused Steel parent row
            

            double virginTotalVolume = 0;
            double virginTotalCarbon = 0;
            double reusedTotalVolume = 0;
            double reusedTotalCarbon = 0;

            // Elements to process
            string[] elementTypes = { "Beams", "Columns", "Floors", "Pilings" };

            // Process each element type
            foreach (var elementType in elementTypes)
            {
                if (!sectionVolumes.ContainsKey(elementType)) continue;

                double elementTotalVolume = 0;
                double elementTotalCarbon = 0;

                // Create element row for Virgin Steel
                var elementRow = new SteelTableRow(_steelData, this)
                {
                    Name = elementType,
                    Level = 1,
                    IsEditable = false,
                    IsManualECFactor = false
                };

                // For Beams and Columns, add section rows
                if (elementType == "Beams" || elementType == "Columns")
                {
                    // Add section rows for this element
                    foreach (var sectionType in new[] { "Open Sections", "Closed Sections", "Plates" })
                    {
                        double volume = sectionVolumes[elementType][sectionType];
                        elementTotalVolume += volume;

                        // Create section row
                        var sectionRow = new SteelTableRow(_steelData, this)
                        {
                            Name = sectionType,
                            Level = 2,
                            IsEditable = true,
                            Volume = volume,
                            Source = "UK", // Default source
                            SectionType = sectionType,
                            IsManualECFactor = false
                        };

                        // Add property changed handler to sync changes back to main model
                        sectionRow.PropertyChanged += SteelTableRow_PropertyChanged;

                        elementTotalCarbon += sectionRow.Carbon;
                        result.Add(sectionRow);
                    }
                }
                else // For Floors and Piles, add a single row
                {
                    double totalVolume = sectionVolumes[elementType].Values.Sum();
                    elementTotalVolume = totalVolume;

                    // Create combined row
                    var combinedRow = new SteelTableRow(_steelData, this)
                    {
                        Name = elementType,
                        Level = 2,
                        IsEditable = true,
                        Volume = totalVolume,
                        Source = "UK", // Default source
                        SectionType = "Open Sections", // Default section type
                        IsManualECFactor = false
                    };

                    // Add property changed handler to sync changes back to main model
                    combinedRow.PropertyChanged += SteelTableRow_PropertyChanged;

                    elementTotalCarbon += combinedRow.Carbon;
                    result.Add(combinedRow);
                }

                // Set the element row's carbon and volume totals
                elementRow.Volume = elementTotalVolume;
                elementRow.SetCarbon(elementTotalCarbon);

                // Insert element row at the correct position
                int insertIndex = result.Count - (elementType == "Beams" || elementType == "Columns" ? 3 : 1);
                result.Insert(insertIndex, elementRow);

                // Add to virgin steel totals
                virginTotalVolume += elementTotalVolume;
                virginTotalCarbon += elementTotalCarbon;
            }

            // Set Virgin Steel totals
            virginSteel.Volume = virginTotalVolume;
            virginSteel.SetCarbon(virginTotalCarbon);

            // Insert the parent rows
            result.Insert(0, virginSteel);
            

            return result;
        }


        public Dictionary<string, Dictionary<string, double>> GetSteelData()
        {
            return _steelData;
        }
        private ElementId FindSteelMaterial(Document doc)
        {
            // Get all materials in the document
            var allMaterials = new FilteredElementCollector(doc)
                .OfClass(typeof(Material))
                .Cast<Material>();

            // Log all materials for debugging
            System.Diagnostics.Debug.WriteLine("All materials in document:");
            foreach (var mat in allMaterials)
            {
                System.Diagnostics.Debug.WriteLine($"- {mat.Name}");
            }

            // First try: Look for exact matches with common steel material names
            string[] steelNames = { "Steel", "WAL_Steel", "Structural Steel", "Metal - Steel" };
            foreach (var steelName in steelNames)
            {
                var exactMatch = allMaterials.FirstOrDefault(m => m.Name.Equals(steelName, StringComparison.OrdinalIgnoreCase));
                if (exactMatch != null)
                {
                    System.Diagnostics.Debug.WriteLine($"Found exact match for steel: {exactMatch.Name}");
                    return exactMatch.Id;
                }
            }

            // Second try: Look for partial matches
            string[] steelKeywords = { "steel", "metal", "stl" };
            foreach (var keyword in steelKeywords)
            {
                var partialMatch = allMaterials.FirstOrDefault(m =>
                    m.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
                if (partialMatch != null)
                {
                    System.Diagnostics.Debug.WriteLine($"Found partial match for steel: {partialMatch.Name}");
                    return partialMatch.Id;
                }
            }

            // If no match found, return null
            System.Diagnostics.Debug.WriteLine("No steel material found in document");
            return null;
        }

        // Update the SteelTableRow_PropertyChanged event handler
        private void SteelTableRow_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (sender is SteelTableRow row &&
                (e.PropertyName == nameof(SteelTableRow.Volume) ||
                 e.PropertyName == nameof(SteelTableRow.Source) ||
                 e.PropertyName == nameof(SteelTableRow.CarbonFactor)))
            {
                // Log for debugging
                System.Diagnostics.Debug.WriteLine($"Steel table row property changed: {e.PropertyName}");

                // Force re-sync for all rows to ensure everything is up to date
                SyncSteelTableChanges();

                // Add this line - directly update the total carbon
                UpdateTotalCarbonEmbodiment();
            }
        }

        // Update the SyncSteelTableChanges method to remove Module references
        public void SyncSteelTableChanges()
        {
            // Skip if the flattened table hasn't been created yet
            if (_flattenedSteelTable == null)
                return;

            // Log that we're starting the sync
            System.Diagnostics.Debug.WriteLine("Starting steel table sync...");

            // Track changes to determine if update is needed
            bool madeChanges = false;

            // First pass: calculate total steel volumes by element type from all section types
            var combinedVolumes = new Dictionary<string, double>();
            // Add dictionary to track carbon values by element type
            var combinedCarbon = new Dictionary<string, double>();

            // Process all editable rows in the flattened table
            foreach (var row in _flattenedSteelTable.Where(r => r.IsEditable))
            {
                // Skip if row has 0 volume
                if (row.Volume <= 0)
                    continue;

                // Extract the element type from the parent row
                string elementType = GetElementTypeFromRow(row);
                if (string.IsNullOrEmpty(elementType) || !_items.ContainsKey(elementType))
                    continue;

                // Add to combined volumes for this element type
                if (!combinedVolumes.ContainsKey(elementType))
                {
                    combinedVolumes[elementType] = 0;
                    combinedCarbon[elementType] = 0;
                }

                combinedVolumes[elementType] += row.Volume;
                combinedCarbon[elementType] += row.Carbon; // Sum up carbon values from child rows
            }

            // Second pass: apply the combined volumes and update properties
            foreach (var kvp in combinedVolumes)
            {
                string elementType = kvp.Key;
                double totalVolume = kvp.Value;
                double totalCarbon = combinedCarbon[elementType];

                if (_items.ContainsKey(elementType))
                {
                    // Update if value has changed
                    if (Math.Abs(_items[elementType].SteelVolume - totalVolume) > 0.001)
                    {
                        System.Diagnostics.Debug.WriteLine($"Updating {elementType} steel volume from {_items[elementType].SteelVolume} to {totalVolume}");
                        _items[elementType].SteelVolume = totalVolume;
                        madeChanges = true;
                    }

                    // Force a recalculation of steel carbon in the items dictionary to match the UI
                    // Use explicit Func<> type for C# 7.3 compatibility
                    Func<SteelTableRow, bool> steelRowPredicate = delegate (SteelTableRow row) { return !row.IsEditable && row.Name == elementType; };
                    var elementRow = _flattenedSteelTable.FirstOrDefault(steelRowPredicate);
                    if (elementRow != null)
                    {
                        // Extract CarbonFactor from child rows - use dominant section's values
                        var dominantRow = _flattenedSteelTable
                            .Where(r => r.IsEditable && GetElementTypeFromRow(r) == elementType)
                            .OrderByDescending(r => r.Volume)
                            .FirstOrDefault();

                        if (dominantRow != null)
                        {
                            _items[elementType].SteelSource = dominantRow.Source;

                            // Map section type to steel type
                            string sectionType = dominantRow.SectionType;
                            string steelType = ConvertSectionTypeToSteelType(sectionType);
                            _items[elementType].SteelType = steelType;

                            System.Diagnostics.Debug.WriteLine($"Updated {elementType} steel properties to match UI: Source={dominantRow.Source}, Type={steelType}");
                        }

                        // Also update the UI row
                        elementRow.Volume = totalVolume;
                        elementRow.SetCarbon(totalCarbon);
                        System.Diagnostics.Debug.WriteLine($"Updated parent row {elementType} carbon to {totalCarbon}");
                        madeChanges = true;
                    }
                }
            }

            // Now update properties for each element type based on its sections
            foreach (var row in _flattenedSteelTable.Where(r => r.IsEditable))
            {
                // Extract the element type from the parent row
                string elementType = GetElementTypeFromRow(row);
                if (string.IsNullOrEmpty(elementType) || !_items.ContainsKey(elementType))
                    continue;

                // Get the section type
                string sectionType = row.SectionType;
                string steelType = ConvertSectionTypeToSteelType(sectionType);

                // Update steel type and source if this is the dominant section for this element
                // (based on volume)
                if (_items.ContainsKey(elementType) && row.Volume > 0)
                {
                    // Find the row with the largest volume for this element type
                    var highestVolumeRow = _flattenedSteelTable
                        .Where(r => r.IsEditable && GetElementTypeFromRow(r) == elementType)
                        .OrderByDescending(r => r.Volume)
                        .FirstOrDefault();

                    // If this is the highest volume row for this element, use its properties
                    if (highestVolumeRow == row)
                    {
                        if (_items[elementType].SteelType != steelType)
                        {
                            _items[elementType].SteelType = steelType;
                            madeChanges = true;
                        }

                        if (_items[elementType].SteelSource != row.Source)
                        {
                            _items[elementType].SteelSource = row.Source;
                            madeChanges = true;
                        }
                    }
                }
            }

            // Update the top-level Virgin Steel row
            double virginTotalVolume = 0;
            double virginTotalCarbon = 0;

            // Sum up all element-level volumes and carbon
            foreach (var row in _flattenedSteelTable.Where(r => !r.IsEditable && r.Level == 1))
            {
                virginTotalVolume += row.Volume;
                virginTotalCarbon += row.Carbon;
            }

            // Update the Virgin Steel parent row
            var virginSteel = _flattenedSteelTable.FirstOrDefault(r => r.Level == 0 && r.Name == "Virgin Steel");
            if (virginSteel != null)
            {
                virginSteel.Volume = virginTotalVolume;
                virginSteel.SetCarbon(virginTotalCarbon);
                System.Diagnostics.Debug.WriteLine($"Updated Virgin Steel row: volume={virginTotalVolume}, carbon={virginTotalCarbon}");
                madeChanges = true;
            }

            // Force an update of the total carbon regardless of detected changes
            System.Diagnostics.Debug.WriteLine("Forcing update of total carbon embodiment...");
            UpdateTotalCarbonEmbodiment();

            foreach (var item in _items)
            {
                if (item.Value.SteelVolume > 0) // Only update items that have steel
                {
                    // Force property change for these key properties
                    item.Value.ForcePropertyChange(nameof(ItemData.SteelVolume));
                    item.Value.ForcePropertyChange(nameof(ItemData.SteelCarbon));
                    item.Value.ForcePropertyChange(nameof(ItemData.SubTotalCarbon));
                }
            }

            // Explicitly update total carbon
            UpdateTotalCarbonEmbodiment();

            // Force UI updates for all relevant properties
            OnPropertyChanged(nameof(TotalCarbonEmbodiment));
            OnPropertyChanged(nameof(CarbonRating));
        }



        // Helper method to map from section type to steel type format
        private string ConvertSectionTypeToSteelType(string sectionType)
        {
            if (sectionType == "Open Sections")
                return "Open Section";
            else if (sectionType == "Closed Sections")
                return "Closed Section";
            else
                return "Plates";
        }

        // Helper method to determine element type from row
        private string GetElementTypeFromRow(SteelTableRow row)
        {
            // For section rows, we need to find the parent element row
            if (row.Level == 2)
            {
                // Find the parent row (level 1) that this section belongs to
                int rowIndex = _flattenedSteelTable.IndexOf(row);
                if (rowIndex <= 0) return null;

                // Look backward to find the parent row
                for (int i = rowIndex - 1; i >= 0; i--)
                {
                    if (_flattenedSteelTable[i].Level == 1)
                    {
                        // Found the parent row
                        return _flattenedSteelTable[i].Name;
                    }
                }
            }
            else if (row.Level == 1)
            {
                // This is already an element row
                return row.Name;
            }

            foreach (var item in _items)
            {
                if (item.Value.SteelVolume > 0) // Only update items that have steel
                {
                    // Force property change for these key properties
                    item.Value.ForcePropertyChange(nameof(ItemData.SteelVolume));
                    item.Value.ForcePropertyChange(nameof(ItemData.SteelCarbon));
                    item.Value.ForcePropertyChange(nameof(ItemData.SubTotalCarbon));
                }
            }

            // Explicitly update total carbon
            UpdateTotalCarbonEmbodiment();

            // Force UI updates for all relevant properties
            OnPropertyChanged(nameof(TotalCarbonEmbodiment));
            OnPropertyChanged(nameof(CarbonRating));



            return null;
        }
        public void UpdateSteelCarbonFactorFromDropdown(SteelTableRow row)
        {
            if (row == null) return;

            // Get the steel type and source
            string sectionType = row.SectionType; // Open Sections, Closed Sections, Plates
            string source = row.Source; // UK, Global, UK (Reused), etc.

            // Map section category to steel data keys
            string steelDataKey = sectionType == "Open Sections" ? "Open Section" :
                                  sectionType == "Closed Sections" ? "Closed Section" : "Plates";

            // Look up the carbon factor for this section type and source
            if (_steelData != null && _steelData.ContainsKey(steelDataKey) &&
                _steelData[steelDataKey].ContainsKey(source))
            {
                // Update the carbon factor based on the selected type and source
                double newFactor = _steelData[steelDataKey][source];

                // Only update if the value has changed
                if (Math.Abs(row.CarbonFactor - newFactor) > 0.0001)
                {
                    row.CarbonFactor = newFactor;
                    System.Diagnostics.Debug.WriteLine($"Updated carbon factor to {newFactor} for {sectionType}, {source}");
                }
            }

            // Make sure all totals are updated
            SyncSteelTableChanges();
        }


        // Add these classes and methods to MainViewModel.cs similar to the steel table structure

        // 1. Create a MasonryTableRow class similar to SteelTableRow
        public class MasonryTableRow : INotifyPropertyChanged
        {
            // Basic properties
            private string _name;
            public string Name
            {
                get => _name;
                set
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }

            private int _level;
            public int Level
            {
                get => _level;
                set
                {
                    _level = value;
                    OnPropertyChanged(nameof(Level));
                    OnPropertyChanged(nameof(IndentMargin));
                    OnPropertyChanged(nameof(FontWeight));
                }
            }

            private double _volume;
            public double Volume
            {
                get => _volume;
                set
                {
                    _volume = value;
                    OnPropertyChanged(nameof(Volume));
                    OnPropertyChanged(nameof(Carbon));
                }
            }

            public void SetCarbonFactorDirectly(double value, bool preserveManualFlag)
            {
                // Store the current manual flag state
                bool currentManualState = IsManualECFactor;

                // Set the carbon factor value directly
                _carbonFactor = value;

                // If we need to preserve the manual flag, restore it
                if (preserveManualFlag)
                {
                    IsManualECFactor = currentManualState;
                }

                // Notify that the properties have changed
                OnPropertyChanged(nameof(CarbonFactor));
                OnPropertyChanged(nameof(Carbon));
            }

            private double _carbonFactor;
            public double CarbonFactor
            {
                get => _carbonFactor;
                set
                {
                    if (Math.Abs(_carbonFactor - value) > 0.0001)
                    {
                        double oldValue = _carbonFactor;
                        _carbonFactor = value;
                        System.Diagnostics.Debug.WriteLine($"CarbonFactor changed for {Name} from {oldValue} to {value}");

                        // Flag to indicate this was manually set
                        IsManualECFactor = true;

                        // Notify dependent properties
                        OnPropertyChanged(nameof(CarbonFactor));
                        OnPropertyChanged(nameof(Carbon));
                    }
                }
            }

            private bool _isManualECFactor = false;
            public bool IsManualECFactor
            {
                get => _isManualECFactor;
                set
                {
                    _isManualECFactor = value;
                    OnPropertyChanged(nameof(IsManualECFactor));
                }
            }

            private bool _isEditable = false;
            public bool IsEditable
            {
                get => _isEditable;
                set
                {
                    _isEditable = value;
                    OnPropertyChanged(nameof(IsEditable));
                    OnPropertyChanged(nameof(EditableVisibility));
                    OnPropertyChanged(nameof(VolumeVisibility));
                }
            }

            // Material-specific properties
            private string _masonryType;
            public string MasonryType
            {
                get => _masonryType;
                set
                {
                    if (_masonryType != value)
                    {
                        _masonryType = value;
                        OnPropertyChanged(nameof(MasonryType));
                        UpdateCarbonFactor();
                    }
                }
            }

            private double _density;
            public double Density
            {
                get => _density;
                set
                {
                    _density = value;
                    OnPropertyChanged(nameof(Density));
                    OnPropertyChanged(nameof(Carbon));
                }
            }

            // Reference to masonry data
            private Dictionary<string, Dictionary<string, double>> _masonryData;

            // Constructor
            public MasonryTableRow(Dictionary<string, Dictionary<string, double>> masonryData = null)
            {
                _masonryData = masonryData;

                // Set default density based on type
                if (_masonryType == "Blockwork")
                    _density = 2000;
                else if (_masonryType == "Brickwork")
                    _density = 1910;
            }

            // Update carbon factor based on masonry type
            private void UpdateCarbonFactor()
            {
                if (IsManualECFactor) return;

                if (_masonryData != null && !string.IsNullOrEmpty(_masonryType))
                {
                    if (_masonryData.ContainsKey(_masonryType) && _masonryData[_masonryType].ContainsKey("A1-A3"))
                    {
                        CarbonFactor = _masonryData[_masonryType]["A1-A3"];
                        IsManualECFactor = false; // Reset flag since we're using default values
                    }
                }

                // Update density based on type
                if (_masonryType == "Blockwork")
                    Density = 2000;
                else if (_masonryType == "Brickwork")
                    Density = 1910;
            }

            // Calculated properties
            public Thickness IndentMargin => new Thickness(Level * 20, 0, 0, 0);

            public System.Windows.FontWeight FontWeight
            {
                get
                {
                    if (Level == 0) return System.Windows.FontWeights.Bold;
                    if (Level == 1) return System.Windows.FontWeights.SemiBold;
                    return System.Windows.FontWeights.Normal;
                }
            }

            public double Carbon
            {
                get
                {
                    if (IsEditable)
                        return Volume * CarbonFactor * Density;

                    // For parent rows (non-editable), the Carbon is calculated in the view model
                    return _carbon;
                }
            }

            private double _carbon;
            public void SetCarbon(double value)
            {
                _carbon = value;
                OnPropertyChanged(nameof(Carbon));
            }

            // Visibility properties for UI binding
            public System.Windows.Visibility EditableVisibility => IsEditable ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            public System.Windows.Visibility VolumeVisibility => IsEditable ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;

            // INotifyPropertyChanged implementation
            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // 2. Add a property for the flattened masonry table
        private ObservableCollection<MasonryTableRow> _flattenedMasonryTable;
        public ObservableCollection<MasonryTableRow> FlattenedMasonryTable
        {
            get
            {
                if (_flattenedMasonryTable == null)
                {
                    _flattenedMasonryTable = CreateFlattenedMasonryTable();
                }
                return _flattenedMasonryTable;
            }
        }

        // 3. Create method to build the flattened masonry table
        private ObservableCollection<MasonryTableRow> CreateFlattenedMasonryTable()
        {
            var result = new ObservableCollection<MasonryTableRow>();

            // Create parent row for Masonry
            var masonryParent = new MasonryTableRow(_masonryData)
            {
                Name = "Masonry",
                Level = 0,
                IsEditable = false
            };

            // Add wall row
            var wallsRow = new MasonryTableRow(_masonryData)
            {
                Name = "Walls",
                Level = 1,
                IsEditable = false
            };

            // Calculate volumes for brick and block
            double brickVolume = 0;
            double blockVolume = 0;

            // Get wall elements with brick or block material
            if (_items.ContainsKey("Walls"))
            {
                // For now, assign the total masonry volume to be split between brick and block
                double totalMasonryVolume = _items["Walls"].MasonryVolume;

                // This is where you would determine actual volumes from Revit
                // For the demo, let's assume a 50/50 split or use what's in ItemData
                brickVolume = totalMasonryVolume * 0.5;
                blockVolume = totalMasonryVolume * 0.5;
            }

            // Create brick row
            var brickRow = new MasonryTableRow(_masonryData)
            {
                Name = "Brickwork",
                Level = 2,
                IsEditable = true,
                Volume = brickVolume,
                MasonryType = "Brickwork",
                Density = 1910
            };

            // Create block row
            var blockRow = new MasonryTableRow(_masonryData)
            {
                Name = "Blockwork",
                Level = 2,
                IsEditable = true,
                Volume = blockVolume,
                MasonryType = "Blockwork",
                Density = 2000
            };

            // Add event handlers for property changes
            brickRow.PropertyChanged += MasonryTableRow_PropertyChanged;
            blockRow.PropertyChanged += MasonryTableRow_PropertyChanged;

            // Calculate totals
            double wallsTotalVolume = brickVolume + blockVolume;
            double wallsTotalCarbon = brickRow.Carbon + blockRow.Carbon;

            // Set wall totals
            wallsRow.Volume = wallsTotalVolume;
            wallsRow.SetCarbon(wallsTotalCarbon);

            // Set masonry parent totals
            masonryParent.Volume = wallsTotalVolume;
            masonryParent.SetCarbon(wallsTotalCarbon);

            // Add rows to result
            result.Add(masonryParent);
            result.Add(wallsRow);
            result.Add(brickRow);
            result.Add(blockRow);

            return result;
        }

        // 4. Add event handler for property changes
        private void MasonryTableRow_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (sender is MasonryTableRow row &&
                (e.PropertyName == nameof(MasonryTableRow.Volume) ||
                 e.PropertyName == nameof(MasonryTableRow.CarbonFactor)))
            {
                System.Diagnostics.Debug.WriteLine($"Masonry table row property changed: {e.PropertyName}");
                SyncMasonryTableChanges();
                UpdateTotalCarbonEmbodiment();
            }
        }

        // 5. Add method to sync changes back to data model
        public void SyncMasonryTableChanges()
        {
            // Skip if table hasn't been created yet
            if (_flattenedMasonryTable == null)
                return;

            System.Diagnostics.Debug.WriteLine("Starting masonry table sync...");

            // Calculate total volume and carbon for walls
            double totalVolume = 0;
            double totalCarbon = 0;

            // Process all editable rows
            foreach (var row in _flattenedMasonryTable.Where(r => r.IsEditable))
            {
                totalVolume += row.Volume;
                totalCarbon += row.Carbon;
            }

            // Update the Walls row
            var wallsRow = _flattenedMasonryTable.FirstOrDefault(r => r.Level == 1 && r.Name == "Walls");
            if (wallsRow != null)
            {
                wallsRow.Volume = totalVolume;
                wallsRow.SetCarbon(totalCarbon);
            }

            // Update the Masonry parent row
            var masonryParent = _flattenedMasonryTable.FirstOrDefault(r => r.Level == 0 && r.Name == "Masonry");
            if (masonryParent != null)
            {
                masonryParent.Volume = totalVolume;
                masonryParent.SetCarbon(totalCarbon);
            }

            // Update the ItemData model
            if (_items.ContainsKey("Walls"))
            {
                _items["Walls"].MasonryVolume = totalVolume;

                // Force notification of property changes
                _items["Walls"].NotifyPropertyChanged(nameof(ItemData.MasonryVolume));
                _items["Walls"].NotifyPropertyChanged(nameof(ItemData.MasonryCarbon));
                _items["Walls"].NotifyPropertyChanged(nameof(ItemData.SubTotalCarbon));
            }

            // Update total carbon
            UpdateTotalCarbonEmbodiment();
        }

        // 6. Add method to update masonry carbon factor from popup
        public void UpdateMasonryCarbonFactorFromPopup(MasonryTableRow row, double newValue)
        {
            if (row == null) return;

            // Only update if the value changed
            if (Math.Abs(row.CarbonFactor - newValue) > 0.0001)
            {
                row.CarbonFactor = newValue;
                row.IsManualECFactor = true;
                System.Diagnostics.Debug.WriteLine($"Updated masonry carbon factor to {newValue} for {row.Name}");
            }

            // Update all totals
            SyncMasonryTableChanges();
        }






    }
}