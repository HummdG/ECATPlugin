using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;

namespace ECATPlugin
{
    public partial class SteelECPopupWindow : UserControl
    {
        public event EventHandler<string> ValueSelected;
        private DataGrid _valueDataGrid;
        private string _sectionType;

        public SteelECPopupWindow(string sectionType)
        {
            _sectionType = sectionType;
            InitializeComponent();
            SetupUI();
        }

        private void SetupUI()
        {
            // Set fixed size for popup
            this.Width = 500;
            this.Height = 250;

            // Set background and border
            this.Background = Brushes.White;
            Border border = new Border
            {
                BorderBrush = Brushes.Gray,
                BorderThickness = new Thickness(1),
                Child = new Grid(),
                Effect = new DropShadowEffect
                {
                    BlurRadius = 10,
                    ShadowDepth = 5,
                    Opacity = 0.3
                }
            };

            // Create the DataGrid
            _valueDataGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                HeadersVisibility = DataGridHeadersVisibility.Column,
                GridLinesVisibility = DataGridGridLinesVisibility.All,
                HorizontalGridLinesBrush = Brushes.LightGray,
                VerticalGridLinesBrush = Brushes.LightGray,
                Background = Brushes.White,
                BorderThickness = new Thickness(0),
                Margin = new Thickness(5),
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                SelectionMode = DataGridSelectionMode.Single,
                SelectionUnit = DataGridSelectionUnit.Cell,
                CanUserSortColumns = false // Disable sorting
            };

            // Add columns to the DataGrid
            _valueDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Source",
                Binding = new System.Windows.Data.Binding("Source"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                CanUserSort = false
            });

            _valueDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Module A1-A3",
                Binding = new System.Windows.Data.Binding("ModuleA1A3"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                CanUserSort = false
            });

            _valueDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Module A4",
                Binding = new System.Windows.Data.Binding("ModuleA4"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                CanUserSort = false
            });

            _valueDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Module A5",
                Binding = new System.Windows.Data.Binding("ModuleA5"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                CanUserSort = false
            });

            // Create style for the source column
            Style sourceColumnStyle = new Style(typeof(DataGridCell));
            sourceColumnStyle.Setters.Add(new Setter(DataGridCell.BackgroundProperty, Brushes.LightGray));
            sourceColumnStyle.Setters.Add(new Setter(DataGridCell.ForegroundProperty, Brushes.Black));
            sourceColumnStyle.Setters.Add(new Setter(DataGridCell.FontWeightProperty, FontWeights.Bold));
            sourceColumnStyle.Setters.Add(new Setter(DataGridCell.BorderBrushProperty, Brushes.Gray));
            sourceColumnStyle.Setters.Add(new Setter(DataGridCell.BorderThicknessProperty, new Thickness(1)));
            sourceColumnStyle.Setters.Add(new Setter(DataGridCell.PaddingProperty, new Thickness(5)));
            sourceColumnStyle.Setters.Add(new Setter(DataGridCell.CursorProperty, Cursors.Arrow)); // Regular cursor

            // Style for the value columns (modules)
            Style valueColumnStyle = new Style(typeof(DataGridCell));
            valueColumnStyle.Setters.Add(new Setter(DataGridCell.BackgroundProperty, Brushes.White));
            valueColumnStyle.Setters.Add(new Setter(DataGridCell.BorderBrushProperty, Brushes.LightGray));
            valueColumnStyle.Setters.Add(new Setter(DataGridCell.BorderThicknessProperty, new Thickness(1)));
            valueColumnStyle.Setters.Add(new Setter(DataGridCell.PaddingProperty, new Thickness(5)));
            valueColumnStyle.Setters.Add(new Setter(DataGridCell.CursorProperty, Cursors.Hand)); // Hand cursor to indicate clickable

            // Add hover effect for value columns
            Trigger mouseOverTrigger = new Trigger { Property = DataGridCell.IsMouseOverProperty, Value = true };
            mouseOverTrigger.Setters.Add(new Setter(DataGridCell.BackgroundProperty, Brushes.LightBlue));
            valueColumnStyle.Triggers.Add(mouseOverTrigger);

            // Apply column styles
            _valueDataGrid.Columns[0].CellStyle = sourceColumnStyle; // Source column
            _valueDataGrid.Columns[1].CellStyle = valueColumnStyle;  // Module A1-A3
            _valueDataGrid.Columns[2].CellStyle = valueColumnStyle;  // Module A4
            _valueDataGrid.Columns[3].CellStyle = valueColumnStyle;  // Module A5

            // Add event handlers
            _valueDataGrid.MouseLeftButtonUp += ValueDataGrid_MouseLeftButtonUp;
            _valueDataGrid.PreviewKeyDown += ValueDataGrid_PreviewKeyDown;

            // Create Grid to hold DataGrid
            var grid = (Grid)border.Child;
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            // Add DataGrid to the Grid
            Grid.SetRow(_valueDataGrid, 0);
            grid.Children.Add(_valueDataGrid);

            // Add title
            TextBlock headerText = new TextBlock
            {
                Text = $"{_sectionType} Embodied Carbon Values (kgCO₂e/kg)",
                FontWeight = FontWeights.Bold,
                FontSize = 14,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 5, 0, 5)
            };

            // Create main layout
            Grid mainGrid = new Grid();
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            Grid.SetRow(headerText, 0);
            Grid.SetRow(border, 1);

            mainGrid.Children.Add(headerText);
            mainGrid.Children.Add(border);

            // Set content
            this.Content = mainGrid;

            // Load data based on section type
            LoadData();
        }

        private void LoadData()
        {
            // Get data based on the section type
            List<SteelECData> data = SteelECDataProvider.GetDataForSectionType(_sectionType);
            _valueDataGrid.ItemsSource = data;
        }

        private void ValueDataGrid_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            // Get the cell that was clicked
            var cell = GetCellUnderMouse(e.GetPosition(_valueDataGrid));
            if (cell != null)
            {
                SelectCellValue(cell);
            }
        }

        private void ValueDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            // Handle Enter key
            if (e.Key == Key.Enter)
            {
                var currentCell = _valueDataGrid.CurrentCell;
                if (currentCell.Column != null)
                {
                    SelectCellValue(currentCell);
                    e.Handled = true;
                }
            }
        }

        private void SelectCellValue(DataGridCellInfo cellInfo)
        {
            if (cellInfo.Column == null) return;

            // Get the header name and item
            string headerName = cellInfo.Column.Header.ToString();
            SteelECData item = cellInfo.Item as SteelECData;

            // Only allow selecting cells from module columns, not from the source column
            if (item != null && headerName != "Source")
            {
                string value = "";

                if (headerName == "Module A1-A3")
                    value = item.ModuleA1A3;
                else if (headerName == "Module A4")
                    value = item.ModuleA4;
                else if (headerName == "Module A5")
                    value = item.ModuleA5;

                if (!string.IsNullOrEmpty(value))
                {
                    ValueSelected?.Invoke(this, value);
                }
            }
        }

        private DataGridCellInfo GetCellUnderMouse(Point mousePosition)
        {
            // Find the visual element under the mouse
            HitTestResult result = VisualTreeHelper.HitTest(_valueDataGrid, mousePosition);
            if (result == null) return new DataGridCellInfo();

            // Traverse up to find the DataGridCell
            DependencyObject element = result.VisualHit;
            while (element != null && !(element is DataGridCell) && !(element is DataGridRow))
            {
                element = VisualTreeHelper.GetParent(element);
            }

            // If we found a cell
            if (element is DataGridCell cell)
            {
                // Skip if this is the source column (column index 0)
                if (cell.Column != null && cell.Column.DisplayIndex == 0)
                    return new DataGridCellInfo();

                // Find the row that contains this cell
                DataGridRow row = FindVisualParent<DataGridRow>(cell);
                if (row != null)
                {
                    int columnIndex = cell.Column.DisplayIndex;
                    return new DataGridCellInfo(row.Item, _valueDataGrid.Columns[columnIndex]);
                }
            }
            // If we only found a row, get the cell based on the X position
            else if (element is DataGridRow row)
            {
                // Get the cell based on the X position
                double accumulatedWidth = 0;
                for (int i = 0; i < _valueDataGrid.Columns.Count; i++)
                {
                    var column = _valueDataGrid.Columns[i];
                    accumulatedWidth += column.ActualWidth;

                    // Skip the source column (column index 0)
                    if (i == 0)
                        continue;

                    if (mousePosition.X < accumulatedWidth)
                    {
                        return new DataGridCellInfo(row.Item, column);
                    }
                }
            }

            return new DataGridCellInfo();
        }

        // Helper method to find a parent of a specific type in the visual tree
        private static T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            // Get parent item
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);

            // We've reached the end of the tree
            if (parentObject == null) return null;

            // Check if the parent matches the type we're looking for
            T parent = parentObject as T;
            if (parent != null)
                return parent;
            else
                return FindVisualParent<T>(parentObject);
        }
    }
}