using System.Collections.Generic;
using ECATPlugin.Helpers;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Controls;
using System.IO;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace ECATPlugin
{
    public partial class MainWindow : Window
    {


        // Add a new popup for steel EC values
        private Popup _steelECPopup;
        private SteelECPopupWindow _steelECPopupContent;
        private TextBox _activeECTextBox = null;
        private MainViewModel.SteelTableRow _activeSteelRow = null;

        private Popup _timberECPopup;
        private TimberECPopupWindow _timberECPopupContent;
        private TextBox _activeTimberECTextBox = null;
        private string _activeTimberType = null;

        private Popup _masonryECPopup;
        private MasonryECPopupWindow _masonryECPopupContent;
        private TextBox _activeMasonryECTextBox = null;
        private MainViewModel.MasonryTableRow _activeMasonryRow = null;

        private bool _isPopupOpen = false;
        private IEnumerable<ConcreteData> _values = new List<ConcreteData>
        {
            new ConcreteData { ConcreteType = "C16/20", ZeroPercent = "275.0", TwentyFivePercent = "205.0", FiftyPercent = "155.0", SeventyFivePercent = "115.0" },
            new ConcreteData { ConcreteType = "C20/25", ZeroPercent = "295.0", TwentyFivePercent = "225.0", FiftyPercent = "165.0", SeventyFivePercent = "120.0" },
            new ConcreteData { ConcreteType = "C25/30", ZeroPercent = "310.0", TwentyFivePercent = "240.0", FiftyPercent = "175.0", SeventyFivePercent = "125.0" },
            new ConcreteData { ConcreteType = "C28/35", ZeroPercent = "330.0", TwentyFivePercent = "260.0", FiftyPercent = "190.0", SeventyFivePercent = "130.0" },
            new ConcreteData { ConcreteType = "C30/37", ZeroPercent = "345.0", TwentyFivePercent = "275.0", FiftyPercent = "205.0", SeventyFivePercent = "135.0" },
            new ConcreteData { ConcreteType = "C32/40", ZeroPercent = "360.0", TwentyFivePercent = "290.0", FiftyPercent = "215.0", SeventyFivePercent = "140.0" },
            new ConcreteData { ConcreteType = "C35/45", ZeroPercent = "390.0", TwentyFivePercent = "310.0", FiftyPercent = "230.0", SeventyFivePercent = "150.0" },
            new ConcreteData { ConcreteType = "C40/50", ZeroPercent = "415.0", TwentyFivePercent = "335.0", FiftyPercent = "245.0", SeventyFivePercent = "155.0" },
            new ConcreteData { ConcreteType = "C50/60", ZeroPercent = "470.0", TwentyFivePercent = "380.0", FiftyPercent = "280.0", SeventyFivePercent = "175.0" },
            new ConcreteData { ConcreteType = "C60/75", ZeroPercent = "520.0", TwentyFivePercent = "425.0", FiftyPercent = "315.0", SeventyFivePercent = "195.0" },
        };

        private MainViewModel _viewModel;

        public MainWindow(MainViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;

            // Set window to be non-modal to allow interaction with Revit
            this.ShowInTaskbar = true;

            InitializePopup();
            InitializeMaterialTables(); // Add this line to initialize material tables

            // Set both view models as resources
            this.Resources["MainViewModel"] = _viewModel;

            // Set the main data context
            DataContext = _viewModel;

            // Hook up event handlers for GIA radiobuttons if they exist in the XAML
            if (InputGIARadioButton != null)
                InputGIARadioButton.Checked += (s, e) => _viewModel.IsInputGIASelected = true;
            if (CalculatedGIARadioButton != null)
                CalculatedGIARadioButton.Checked += (s, e) => _viewModel.IsCalculatedGIASelected = true;

            // Handle window closing to properly save data
            this.Closing += (s, e) => _viewModel.Dispose();

            // Register the TextBox event handlers
            RegisterECTextBoxHandlers();
        }

        private void RegisterECTextBoxHandlers()
        {
            // Find all TextBoxes in the window and register event handlers for them
            var textBoxes = FindTextBoxesInVisualTree(this);
            foreach (var textBox in textBoxes)
            {
                // Check if these are the embodied carbon textboxes based on naming conventions
                if (textBox.Name != null && (textBox.Name.Contains("EC") || textBox.Name.Contains("EmbodiedCarbon")))
                {
                    textBox.PreviewMouseDown += EmbodiedCarbonTextBox_PreviewMouseDown;
                    textBox.TextChanged += EmbodiedCarbonTextBox_TextChanged;
                    textBox.PreviewTextInput += NumberOnly_PreviewTextInput;

                    // Remove LostFocus handler as we want the popup to stay open
                    // textBox.LostFocus += EmbodiedCarbonTextBox_LostFocus;
                }
            }
        }

        private List<TextBox> FindTextBoxesInVisualTree(DependencyObject parent)
        {
            var result = new List<TextBox>();

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is TextBox textBox)
                {
                    result.Add(textBox);
                }
                else
                {
                    result.AddRange(FindTextBoxesInVisualTree(child));
                }
            }

            return result;
        }

        private Popup _valuePopup;
        private PopupWindow _popupContent;
        private TextBox _activeTextBox = null;

        private void InitializePopup()
        {
            _valuePopup = new Popup
            {
                Placement = PlacementMode.Bottom,
                StaysOpen = true,
                AllowsTransparency = true
            };

            // Create popup content for concrete
            _popupContent = new PopupWindow();
            _popupContent.SetValues(_values);
            _popupContent.ValueSelected += PopupContent_ValueSelected;

            // Set the popup content
            _valuePopup.Child = _popupContent;
            _valuePopup.Closed += (s, e) => _isPopupOpen = false;

            // Create popup for steel EC values
            _steelECPopup = new Popup
            {
                Placement = PlacementMode.Bottom,
                StaysOpen = true,
                AllowsTransparency = true
            };
            _steelECPopup.Closed += (s, e) => _isPopupOpen = false;

            // Create popup for timber EC values
            _timberECPopup = new Popup
            {
                Placement = PlacementMode.Bottom,
                StaysOpen = true,
                AllowsTransparency = true
            };
            _timberECPopup.Closed += (s, e) => _isPopupOpen = false;

            _masonryECPopup = new Popup
            {
                Placement = PlacementMode.Bottom,
                StaysOpen = true,
                AllowsTransparency = true
            };
            _masonryECPopup.Closed += (s, e) => _isPopupOpen = false;
        }



        private void NumberOnly_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9.-]+");
            e.Handled = regex.IsMatch(e.Text);

            if (_isPopupOpen)
            {
                ClosePopup();
            }
        }

        private void EmbodiedCarbonTextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                _activeTextBox = textBox;
                if (!_isPopupOpen)
                {
                    _valuePopup.PlacementTarget = textBox;
                    _valuePopup.IsOpen = true;
                    _isPopupOpen = true;
                }
            }
        }


        private bool IsMouseOverPopup()
        {
            if (_popupContent == null) return false;

            Point mousePos = Mouse.GetPosition(_popupContent);
            return mousePos.X >= 0 && mousePos.Y >= 0 &&
                   mousePos.X <= _popupContent.ActualWidth &&
                   mousePos.Y <= _popupContent.ActualHeight;
        }

        private void EmbodiedCarbonTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isPopupOpen)
            {
                ClosePopup();
            }
        }



        private void PopupContent_ValueSelected(object sender, string selectedValue)
        {
            if (_activeTextBox != null)
            {
                _activeTextBox.Text = selectedValue;

                // Set focus to the textbox to ensure the binding is updated
                _activeTextBox.Focus();

                // Explicitly force the binding update
                var bindingExpression = _activeTextBox.GetBindingExpression(TextBox.TextProperty);
                if (bindingExpression != null)
                {
                    bindingExpression.UpdateSource();
                }

                // Small delay to ensure binding completes before closing popup
                System.Windows.Threading.DispatcherTimer timer = new System.Windows.Threading.DispatcherTimer();
                timer.Interval = TimeSpan.FromMilliseconds(100);
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    ClosePopup();
                };
                timer.Start();
                return; // Don't close popup immediately
            }
            ClosePopup();
        }



        private void CaptureScreenButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Create a RenderTargetBitmap to render the visual (the window in this case)
                RenderTargetBitmap renderTarget = new RenderTargetBitmap(
                    (int)this.ActualWidth,
                    (int)this.ActualHeight,
                    96, 96, // DPI (dots per inch)
                    PixelFormats.Pbgra32);

                // Render the current window to the bitmap
                renderTarget.Render(this);

                // Encode the bitmap to a PNG image
                PngBitmapEncoder pngEncoder = new PngBitmapEncoder();
                pngEncoder.Frames.Add(BitmapFrame.Create(renderTarget));

                // Construct the file name using project number and project name
                string fileName = $"{_viewModel.ProjectNumber}_{_viewModel.ProjectName}.png";

                // Replace invalid characters in the file name
                fileName = string.Concat(fileName.Split(Path.GetInvalidFileNameChars()));

                // Specify the file path and save the PNG
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
                {
                    pngEncoder.Save(fileStream);
                }

                MessageBox.Show($"Screenshot saved to {filePath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error capturing screenshot: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Material Type selection changed handler
        private void MaterialType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem != null)
            {
                string materialType = comboBox.Tag as string;

                // Get the item's context
                KeyValuePair<string, MainViewModel.ItemData>? itemPair = null;

                if (comboBox.DataContext is KeyValuePair<string, MainViewModel.ItemData> pair)
                {
                    itemPair = pair;
                }

                // Validate source selection is compatible with type
                if (itemPair.HasValue)
                {
                    bool ecReset = false;

                    if (materialType == "Steel")
                    {
                        string steelType = itemPair.Value.Value.SteelType;
                        bool validSource = false;

                        // Get available sources for this steel type
                        var availableSources = _viewModel.GetAvailableSteelSources(steelType);

                        // Check if current source is valid
                        if (availableSources.Contains(itemPair.Value.Value.SteelSource))
                        {
                            validSource = true;
                        }

                        // If not valid, select first available source
                        if (!validSource && availableSources.Count > 0)
                        {
                            itemPair.Value.Value.SteelSource = availableSources[0];
                            ecReset = true;
                        }

                        // Reset EC factor based on the new type selection
                        if (e.AddedItems.Count > 0)
                        {
                            // This is a new selection - reset the carbon factor
                            ecReset = true;
                        }
                    }
                    else if (materialType == "Timber")
                    {
                        string timberType = itemPair.Value.Value.TimberType;
                        bool validSource = false;

                        // Get available sources for this timber type
                        var availableSources = _viewModel.GetAvailableTimberSources(timberType);

                        // Check if current source is valid
                        if (availableSources.Contains(itemPair.Value.Value.TimberSource))
                        {
                            validSource = true;
                        }

                        // If not valid, select first available source
                        if (!validSource && availableSources.Count > 0)
                        {
                            itemPair.Value.Value.TimberSource = availableSources[0];
                            ecReset = true;
                        }

                        // Reset TimberECFactor when type changes and set it to the appropriate value
                        if (e.AddedItems.Count > 0)
                        {
                            // Only update if it hasn't been manually set
                            if (!itemPair.Value.Value.IsManualTimberECFactor)
                            {
                                // Look up the default value for this timber type
                                double defaultTimberECFactor = GetDefaultTimberECFactor(timberType, itemPair.Value.Value.TimberSource);

                                // Set it to the default value instead of 0
                                itemPair.Value.Value.TimberECFactor = defaultTimberECFactor;

                                // This wasn't manually set, so we should reset the flag
                                itemPair.Value.Value.IsManualTimberECFactor = false;
                            }
                            ecReset = true;
                        }
                    }
                    else if (materialType == "Masonry")
                    {
                        // Reset MasonryECFactor when type changes
                        if (e.AddedItems.Count > 0)
                        {
                            // Reset to 0 to use the default calculation
                            itemPair.Value.Value.MasonryECFactor = 0;
                            ecReset = true;
                        }
                    }

                    // If we've made changes that affect EC, make sure UI updates
                    if (ecReset)
                    {
                        // Force property changes to update UI
                        itemPair.Value.Value.NotifyPropertyChanged("EC");
                        if (materialType == "Steel")
                            itemPair.Value.Value.NotifyPropertyChanged("SteelCarbon");
                        else if (materialType == "Timber")
                            itemPair.Value.Value.NotifyPropertyChanged("TimberCarbon");
                        else if (materialType == "Masonry")
                            itemPair.Value.Value.NotifyPropertyChanged("MasonryCarbon");

                        itemPair.Value.Value.NotifyPropertyChanged("SubTotalCarbon");
                    }
                }

                // Update calculations
                _viewModel.UpdateTotalCarbonEmbodiment();
            }
        }

        private void TimberECTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox textBox && textBox.DataContext is KeyValuePair<string, MainViewModel.ItemData> pair)
            {
                // Mark as manually set when user types in the textbox
                pair.Value.IsManualTimberECFactor = true;
            }

            if (_isPopupOpen)
            {
                ClosePopup();
            }
        }

        private double GetDefaultTimberECFactor(string timberType, string timberSource)
        {
            // Values from TimberECDataProvider or hardcoded based on your model values
            if (timberType == "Softwood")
            {
                return 0.263; // Default value for Softwood
            }
            else if (timberType == "Glulam")
            {
                return timberSource == "UK & EU" ? 0.280 : 0.512;
            }
            else if (timberType == "LVL")
            {
                return 0.390;
            }
            else if (timberType == "CLT")
            {
                return 0.420;
            }

            return 0.263; // Default to Softwood if no match
        }

        // Source and Module selection for materials
        private void MaterialSource_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem != null)
            {
                // Get the material type from the tag
                string materialType = comboBox.Tag as string;

                // Get the item's context
                if (comboBox.DataContext is KeyValuePair<string, MainViewModel.ItemData> pair)
                {
                    // Reset EC factors when source changes
                    if (materialType == "Steel")
                    {
                        // Steel source affects carbon factor automatically
                        pair.Value.NotifyPropertyChanged("SteelCarbon");
                    }
                    else if (materialType == "Timber")
                    {
                        // Reset TimberECFactor to 0 to use the default calculation
                        pair.Value.TimberECFactor = 0;
                        pair.Value.NotifyPropertyChanged("TimberCarbon");
                    }
                    else if (materialType == "Masonry")
                    {
                        // Reset MasonryECFactor to 0 to use the default calculation
                        pair.Value.MasonryECFactor = 0;
                        pair.Value.NotifyPropertyChanged("MasonryCarbon");
                    }

                    // Update the subtotal carbon
                    pair.Value.NotifyPropertyChanged("SubTotalCarbon");
                }

                // Trigger property change notifications in the view model
                _viewModel.UpdateTotalCarbonEmbodiment();
            }
        }

        private void MaterialModule_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem != null)
            {
                // Get the material type from the tag
                string materialType = comboBox.Tag as string;

                // Trigger recalculation for traditional data model
                _viewModel.UpdateTotalCarbonEmbodiment();
            }
        }

        // Add these methods to MainWindow.xaml.cs

        // Initialize material tables
        private void InitializeMaterialTables()
        {
            // Register event handlers for material-specific tables
            RegisterMaterialTableEventHandlers();
        }

        private void RegisterMaterialTableEventHandlers()
        {
            // Find each DataGrid in the GroupBox headers for materials
            var steelGrid = FindVisualChildren<DataGrid>(this)
                .FirstOrDefault(grid => grid.Parent is GroupBox gb && gb.Header?.ToString() == "Steel Components");

            var timberGrid = FindVisualChildren<DataGrid>(this)
                .FirstOrDefault(grid => grid.Parent is GroupBox gb && gb.Header?.ToString() == "Timber Components");

            var masonryGrid = FindVisualChildren<DataGrid>(this)
                .FirstOrDefault(grid => grid.Parent is GroupBox gb && gb.Header?.ToString() == "Masonry Components");

            // Add event handlers for DataGrids if found
            if (steelGrid != null)
            {
                steelGrid.Loaded += (s, e) =>
                {
                    // Verify or update correct data source is used
                    if (steelGrid.ItemsSource != _viewModel.SteelItemsView)
                        steelGrid.ItemsSource = _viewModel.SteelItemsView;
                };
            }

            if (timberGrid != null)
            {
                timberGrid.Loaded += (s, e) =>
                {
                    // Verify or update correct data source is used
                    if (timberGrid.ItemsSource != _viewModel.TimberItemsView)
                        timberGrid.ItemsSource = _viewModel.TimberItemsView;
                };
            }

            if (masonryGrid != null)
            {
                masonryGrid.Loaded += (s, e) =>
                {
                    // Verify or update correct data source is used
                    if (masonryGrid.ItemsSource != _viewModel.MasonryItemsView)
                        masonryGrid.ItemsSource = _viewModel.MasonryItemsView;
                };
            }
        }

        // Helper method to find visual children of a specific type
        private static IEnumerable<T> FindVisualChildren<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) yield break;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child != null)
                {
                    if (child is T childOfType)
                        yield return childOfType;

                    foreach (var grandChild in FindVisualChildren<T>(child))
                        yield return grandChild;
                }
            }
        }

        // 1. First, update the RefreshButton_Click method in MainWindow.xaml.cs

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait; // Set wait cursor

                // Temporarily store user-edited values before refresh
                Dictionary<string, Dictionary<string, object>> userEdits = CaptureUserEdits();

                // Refresh volumes and other data from the Revit model
                _viewModel.RefreshVolumes();

                // Restore user edits where appropriate
                RestoreUserEdits(userEdits);

                // Show success message
                MessageBox.Show("Refresh completed successfully. Model data has been updated.",
                               "Refresh Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during refresh: {ex.Message}",
                               "Refresh Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null; // Reset cursor
            }
        }

        // 2. Add helper methods to capture and restore user edits

        private Dictionary<string, Dictionary<string, object>> CaptureUserEdits()
        {
            var result = new Dictionary<string, Dictionary<string, object>>();

            // Capture EC values and other user editable fields from concrete
            foreach (var item in _viewModel.Items)
            {
                var itemEdits = new Dictionary<string, object>();
                itemEdits["EC"] = item.Value.EC;
                itemEdits["RebarDensity"] = item.Value.RebarDensity;

                // Capture manual flags
                if (item.Value.TimberECFactor > 0 && item.Value.IsManualTimberECFactor)
                {
                    itemEdits["TimberECFactor"] = item.Value.TimberECFactor;
                    itemEdits["IsManualTimberECFactor"] = true;
                }

                if (item.Value.MasonryECFactor > 0)
                {
                    itemEdits["MasonryECFactor"] = item.Value.MasonryECFactor;
                    // Add masonry manual flag if exists
                }

                result[item.Key] = itemEdits;
            }

            // Capture Steel table custom EC values
            if (_viewModel.FlattenedSteelTable != null)
            {
                var steelEdits = new Dictionary<string, object>();

                foreach (var row in _viewModel.FlattenedSteelTable.Where(r => r.IsEditable && r.IsManualECFactor))
                {
                    string key = $"Steel_{row.Name}_{row.SectionType}";
                    steelEdits[key] = row.CarbonFactor;
                }

                result["SteelTable"] = steelEdits;
            }

            // Capture Masonry table custom EC values
            if (_viewModel.FlattenedMasonryTable != null)
            {
                var masonryEdits = new Dictionary<string, object>();

                foreach (var row in _viewModel.FlattenedMasonryTable.Where(r => r.IsEditable && r.IsManualECFactor))
                {
                    string key = $"Masonry_{row.Name}";
                    masonryEdits[key] = row.CarbonFactor;
                }

                result["MasonryTable"] = masonryEdits;
            }

            return result;
        }

        private void RestoreUserEdits(Dictionary<string, Dictionary<string, object>> userEdits)
        {
            // Restore concrete values
            foreach (var item in _viewModel.Items)
            {
                if (userEdits.ContainsKey(item.Key))
                {
                    var itemEdits = userEdits[item.Key];

                    // Restore EC values
                    if (itemEdits.ContainsKey("EC"))
                        item.Value.EC = (double)itemEdits["EC"];

                    // Restore rebar density
                    if (itemEdits.ContainsKey("RebarDensity"))
                        item.Value.RebarDensity = (double)itemEdits["RebarDensity"];

                    // Restore timber EC factors
                    if (itemEdits.ContainsKey("TimberECFactor") && itemEdits.ContainsKey("IsManualTimberECFactor"))
                    {
                        item.Value.TimberECFactor = (double)itemEdits["TimberECFactor"];
                        item.Value.IsManualTimberECFactor = (bool)itemEdits["IsManualTimberECFactor"];
                    }

                    // Restore masonry EC factors
                    if (itemEdits.ContainsKey("MasonryECFactor"))
                        item.Value.MasonryECFactor = (double)itemEdits["MasonryECFactor"];

                    // Force property change notifications
                    item.Value.NotifyPropertyChanged("EC");
                    item.Value.NotifyPropertyChanged("RebarDensity");
                    item.Value.NotifyPropertyChanged("TimberECFactor");
                    item.Value.NotifyPropertyChanged("MasonryECFactor");
                    item.Value.NotifyPropertyChanged("SteelCarbon");
                    item.Value.NotifyPropertyChanged("TimberCarbon");
                    item.Value.NotifyPropertyChanged("MasonryCarbon");
                    item.Value.NotifyPropertyChanged("SubTotalCarbon");
                }
            }

            // Restore steel table values
            if (userEdits.ContainsKey("SteelTable") && _viewModel.FlattenedSteelTable != null)
            {
                var steelEdits = userEdits["SteelTable"];

                foreach (var row in _viewModel.FlattenedSteelTable.Where(r => r.IsEditable))
                {
                    string key = $"Steel_{row.Name}_{row.SectionType}";
                    if (steelEdits.ContainsKey(key))
                    {
                        row.CarbonFactor = (double)steelEdits[key];
                        row.IsManualECFactor = true;
                    }
                }

                // Sync changes back to model
                _viewModel.SyncSteelTableChanges();
            }

            // Restore masonry table values
            if (userEdits.ContainsKey("MasonryTable") && _viewModel.FlattenedMasonryTable != null)
            {
                var masonryEdits = userEdits["MasonryTable"];

                foreach (var row in _viewModel.FlattenedMasonryTable.Where(r => r.IsEditable))
                {
                    string key = $"Masonry_{row.Name}";
                    if (masonryEdits.ContainsKey(key))
                    {
                        row.CarbonFactor = (double)masonryEdits[key];
                        row.IsManualECFactor = true;
                    }
                }

                // Sync changes back to model
                _viewModel.SyncMasonryTableChanges();
            }

            // Update total carbon
            _viewModel.UpdateTotalCarbonEmbodiment();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Save data before closing
                if (_viewModel != null)
                {
                    _viewModel.SaveData();
                }

                // Close the window safely
                this.Close();
            }
            catch (Exception ex)
            {
                // Log error but don't crash
                System.Diagnostics.Debug.WriteLine($"Error closing window: {ex.Message}");
                MessageBox.Show($"Error while closing: {ex.Message}", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void SteelCarbonFactor_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Simple approach - just update the total carbon
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.UpdateTotalCarbonEmbodiment();
            }
        }

        private void SteelCarbonFactor_LostFocus(object sender, RoutedEventArgs e)
        {
            // Not needed anymore - the TextChanged event will handle this
        }

        private void SteelType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem != null &&
                comboBox.DataContext is MainViewModel.SteelTableRow row)
            {
                // Update the carbon factor based on the new selection
                if (_viewModel != null)
                {
                    _viewModel.UpdateSteelCarbonFactorFromDropdown(row);
                }
            }
        }

        // In MainWindow.xaml.cs, update SteelSource_SelectionChanged:
        private void SteelSource_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem != null)
            {
                // Get the row data context (the SteelTableRow)
                var row = comboBox.DataContext as MainViewModel.SteelTableRow;
                if (row == null) return;

                // When source changes from dropdown, we want to update EC values but track that this was from dropdown
                // Get the section type and source
                string sectionType = row.SectionType;
                string source = row.Source;

                // Map section category to steel data keys
                string steelDataKey = sectionType == "Open Sections" ? "Open Section" :
                                    sectionType == "Closed Sections" ? "Closed Section" : "Plates";

                // Get the steel data from the view model
                var steelData = _viewModel.GetSteelData();

                // Look up the carbon factor for this section type and source
                if (steelData != null && steelData.ContainsKey(steelDataKey) &&
                    steelData[steelDataKey].ContainsKey(source))
                {
                    // Update the carbon factor based on the selected type and source
                    double newFactor = steelData[steelDataKey][source];

                    // Set the carbon factor, but mark this as a dropdown-initiated change
                    row.SetCarbonFactorFromDropdown(newFactor);

                    System.Diagnostics.Debug.WriteLine($"Updated carbon factor to {newFactor} for {sectionType}, {source}");
                }

                // Make sure all totals are updated
                _viewModel.SyncSteelTableChanges();
            }
        }

        // Add a new method for handling steel EC textbox mouse down events
        private void SteelECTextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                _activeECTextBox = textBox;

                // Get the data context (SteelTableRow) to determine section type
                _activeSteelRow = textBox.DataContext as MainViewModel.SteelTableRow;

                if (_activeSteelRow != null && !_isPopupOpen)
                {
                    // Create the popup content based on section type
                    _steelECPopupContent = new SteelECPopupWindow(_activeSteelRow.SectionType);
                    _steelECPopupContent.ValueSelected += SteelECPopupContent_ValueSelected;

                    // Set the popup content and show it
                    _steelECPopup.Child = _steelECPopupContent;
                    _steelECPopup.PlacementTarget = textBox;
                    _steelECPopup.IsOpen = true;
                    _isPopupOpen = true;
                }
            }
        }

        // Add a method to handle steel EC value selection
        private void SteelECPopupContent_ValueSelected(object sender, string selectedValue)
        {
            if (_activeECTextBox != null)
            {
                _activeECTextBox.Text = selectedValue;

                // Set focus to ensure binding is updated
                _activeECTextBox.Focus();

                // Force binding update
                var bindingExpression = _activeECTextBox.GetBindingExpression(TextBox.TextProperty);
                if (bindingExpression != null)
                {
                    bindingExpression.UpdateSource();
                }

                // Small delay to ensure binding completes before closing popup
                System.Windows.Threading.DispatcherTimer timer = new System.Windows.Threading.DispatcherTimer();
                timer.Interval = TimeSpan.FromMilliseconds(100);
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    ClosePopup();
                };
                timer.Start();
                return;
            }
            ClosePopup();
        }

        // Modify ClosePopup method to close both popups
        private void ClosePopup()
        {
            _valuePopup.IsOpen = false;
            _steelECPopup.IsOpen = false;
            _timberECPopup.IsOpen = false;
            _masonryECPopup.IsOpen = false;
            _isPopupOpen = false;
        }

        // Update the OnPreviewMouseDown method to include the masonry popup
        protected override void OnPreviewMouseDown(MouseButtonEventArgs e)
        {
            if (_isPopupOpen)
            {
                // Check if click is outside all popups and active textbox
                bool clickedOnConcretePopup = _popupContent != null && _popupContent.IsMouseOver;
                bool clickedOnSteelPopup = _steelECPopupContent != null && _steelECPopupContent.IsMouseOver;
                bool clickedOnTimberPopup = _timberECPopupContent != null && _timberECPopupContent.IsMouseOver;
                bool clickedOnMasonryPopup = _masonryECPopupContent != null && _masonryECPopupContent.IsMouseOver;
                bool clickedOnTextBox = _activeECTextBox != null && _activeECTextBox.IsMouseOver ||
                                        _activeTimberECTextBox != null && _activeTimberECTextBox.IsMouseOver ||
                                        _activeMasonryECTextBox != null && _activeMasonryECTextBox.IsMouseOver;

                if (!clickedOnConcretePopup && !clickedOnSteelPopup && !clickedOnTimberPopup &&
                    !clickedOnMasonryPopup && !clickedOnTextBox)
                {
                    ClosePopup();
                }
            }

            base.OnPreviewMouseDown(e);
        }

        private void TimberECTextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                _activeTimberECTextBox = textBox;

                // Get the data context to determine timber type
                if (textBox.DataContext is KeyValuePair<string, MainViewModel.ItemData> pair)
                {
                    string timberType = pair.Value.TimberType;
                    _activeTimberType = timberType;

                    if (!string.IsNullOrEmpty(_activeTimberType) && !_isPopupOpen)
                    {
                        // Create the popup content based on timber type
                        _timberECPopupContent = new TimberECPopupWindow(_activeTimberType);
                        _timberECPopupContent.ValueSelected += TimberECPopupContent_ValueSelected;

                        // Set the popup content and show it
                        _timberECPopup.Child = _timberECPopupContent;
                        _timberECPopup.PlacementTarget = textBox;
                        _timberECPopup.IsOpen = true;
                        _isPopupOpen = true;
                    }
                }
            }
        }

        // Add method to handle timber EC value selection
        private void TimberECPopupContent_ValueSelected(object sender, string selectedValue)
        {
            if (_activeTimberECTextBox != null)
            {
                _activeTimberECTextBox.Text = selectedValue;

                // Set focus to ensure binding is updated
                _activeTimberECTextBox.Focus();

                // Get the data context to access the item
                if (_activeTimberECTextBox.DataContext is KeyValuePair<string, MainViewModel.ItemData> pair)
                {
                    // After setting the value via the popup, ensure it's marked as manual
                    pair.Value.IsManualTimberECFactor = true;
                }

                // Force binding update
                var bindingExpression = _activeTimberECTextBox.GetBindingExpression(TextBox.TextProperty);
                if (bindingExpression != null)
                {
                    bindingExpression.UpdateSource();
                }

                // Small delay to ensure binding completes before closing popup
                System.Windows.Threading.DispatcherTimer timer = new System.Windows.Threading.DispatcherTimer();
                timer.Interval = TimeSpan.FromMilliseconds(100);
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    ClosePopup();
                };
                timer.Start();
                return;
            }
            ClosePopup();
        }

        private void MasonryECTextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                _activeMasonryECTextBox = textBox;

                // Get the data context (MasonryTableRow) to determine masonry type
                _activeMasonryRow = textBox.DataContext as MainViewModel.MasonryTableRow;

                if (_activeMasonryRow != null && !_isPopupOpen)
                {
                    // Create the popup content based on masonry type
                    _masonryECPopupContent = new MasonryECPopupWindow(_activeMasonryRow.MasonryType);
                    _masonryECPopupContent.ValueSelected += MasonryECPopupContent_ValueSelected;

                    // Set the popup content and show it
                    _masonryECPopup.Child = _masonryECPopupContent;
                    _masonryECPopup.PlacementTarget = textBox;
                    _masonryECPopup.IsOpen = true;
                    _isPopupOpen = true;
                }
            }
        }

        // Add handler for masonry EC value selection
        private void MasonryECPopupContent_ValueSelected(object sender, string selectedValue)
        {
            if (_activeMasonryECTextBox != null && _activeMasonryRow != null)
            {
                // Convert the selected value to double
                if (double.TryParse(selectedValue, out double newValue))
                {
                    // Update the carbon factor in the model
                    _viewModel.UpdateMasonryCarbonFactorFromPopup(_activeMasonryRow, newValue);
                }

                // Update the text box
                _activeMasonryECTextBox.Text = selectedValue;

                // Set focus to ensure binding is updated
                _activeMasonryECTextBox.Focus();

                // Force binding update
                var bindingExpression = _activeMasonryECTextBox.GetBindingExpression(TextBox.TextProperty);
                if (bindingExpression != null)
                {
                    bindingExpression.UpdateSource();
                }

                // Small delay to ensure binding completes before closing popup
                System.Windows.Threading.DispatcherTimer timer = new System.Windows.Threading.DispatcherTimer();
                timer.Interval = TimeSpan.FromMilliseconds(100);
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    ClosePopup();
                };
                timer.Start();
                return;
            }
            ClosePopup();
        }



    }
}