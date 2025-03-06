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

namespace ECATPlugin
{
    public partial class MainWindow : Window
    {
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
            
            // Update hierarchical data from the main view model

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
            // Create popup control
            _valuePopup = new Popup
            {
                Placement = PlacementMode.Bottom,
                StaysOpen = true, // Changed to true to keep popup open
                AllowsTransparency = true
            };

            // Create popup content
            _popupContent = new PopupWindow();
            _popupContent.SetValues(_values);
            _popupContent.ValueSelected += PopupContent_ValueSelected;

            // Set the popup content
            _valuePopup.Child = _popupContent;
            _valuePopup.Closed += (s, e) => _isPopupOpen = false;
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
                // Don't prevent focus from being set on the TextBox
                // e.Handled = true;
            }
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

        protected override void OnPreviewMouseDown(MouseButtonEventArgs e)
        {
            if (_isPopupOpen)
            {
                // If clicking outside both the popup and the active textbox, close the popup
                bool clickedOnPopup = _popupContent != null && _popupContent.IsMouseOver;
                bool clickedOnTextBox = _activeTextBox != null && _activeTextBox.IsMouseOver;

                if (!clickedOnPopup && !clickedOnTextBox)
                {
                    ClosePopup();
                }
            }

            base.OnPreviewMouseDown(e);
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

        private void ClosePopup()
        {
            _valuePopup.IsOpen = false;
            _isPopupOpen = false;
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
                        }
                    }
                }

                // Update calculations
                _viewModel.UpdateTotalCarbonEmbodiment();
                
                // Update hierarchical data

            }
        }

        // Source and Module selection for materials
        private void MaterialSource_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem != null)
            {
                // Trigger property change notifications in the view model
                _viewModel.UpdateTotalCarbonEmbodiment();
                
                // Update hierarchical data

            }
        }

        private void MaterialModule_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem != null)
            {
                // Get the material type from the tag
                string materialType = comboBox.Tag as string;
                
                // Check if we're working with a hierarchical item
                
               
                    // Trigger recalculation for traditional data model
                    _viewModel.UpdateTotalCarbonEmbodiment();
               
                
                // Update hierarchical data
   
            }
        }
        
        


    }
}