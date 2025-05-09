﻿using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace ECATPlugin
{
    public partial class PopupWindow : UserControl
    {
        public event EventHandler<string> ValueSelected;
        private DataGrid _valueDataGrid;

        public PopupWindow()
        {
            InitializeComponent();

            // Set fixed size for popup
            this.Width = 500;
            this.Height = 300;

            // Set background and border
            this.Background = Brushes.White;
            Border border = new Border
            {
                BorderBrush = Brushes.Gray,
                BorderThickness = new Thickness(1),
                Child = new Grid(),
                Effect = new System.Windows.Media.Effects.DropShadowEffect
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


            // Add columns to the DataGrid with correct DataGridLengthUnitType
            _valueDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Concrete Type",
                Binding = new System.Windows.Data.Binding("ConcreteType"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                CanUserSort = false // Disable sorting for this column
            });

            _valueDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "0%",
                Binding = new System.Windows.Data.Binding("ZeroPercent"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                CanUserSort = false // Disable sorting for this column
            });

            _valueDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "25%",
                Binding = new System.Windows.Data.Binding("TwentyFivePercent"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                CanUserSort = false // Disable sorting for this column
            });

            _valueDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "50%",
                Binding = new System.Windows.Data.Binding("FiftyPercent"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                CanUserSort = false // Disable sorting for this column
            });

            _valueDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "75%",
                Binding = new System.Windows.Data.Binding("SeventyFivePercent"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                CanUserSort = false // Disable sorting for this column
            });

            // Create different styles for the concrete type column versus the value columns
            Style concreteTypeColumnStyle = new Style(typeof(DataGridCell));
            concreteTypeColumnStyle.Setters.Add(new Setter(DataGridCell.BackgroundProperty, Brushes.LightGray));
            concreteTypeColumnStyle.Setters.Add(new Setter(DataGridCell.ForegroundProperty, Brushes.Black));
            concreteTypeColumnStyle.Setters.Add(new Setter(DataGridCell.FontWeightProperty, FontWeights.Bold));
            concreteTypeColumnStyle.Setters.Add(new Setter(DataGridCell.BorderBrushProperty, Brushes.Gray));
            concreteTypeColumnStyle.Setters.Add(new Setter(DataGridCell.BorderThicknessProperty, new Thickness(1)));
            concreteTypeColumnStyle.Setters.Add(new Setter(DataGridCell.PaddingProperty, new Thickness(5)));
            concreteTypeColumnStyle.Setters.Add(new Setter(DataGridCell.CursorProperty, Cursors.Arrow)); // Regular arrow cursor instead of hand

            // Style the cells to make it clear they are selectable
            Style valueColumnStyle = new Style(typeof(DataGridCell));
            valueColumnStyle.Setters.Add(new Setter(DataGridCell.BackgroundProperty, Brushes.White));
            valueColumnStyle.Setters.Add(new Setter(DataGridCell.BorderBrushProperty, Brushes.LightGray));
            valueColumnStyle.Setters.Add(new Setter(DataGridCell.BorderThicknessProperty, new Thickness(1)));
            valueColumnStyle.Setters.Add(new Setter(DataGridCell.PaddingProperty, new Thickness(5)));
            valueColumnStyle.Setters.Add(new Setter(DataGridCell.CursorProperty, Cursors.Hand)); // Hand cursor to indicate clickable

            // Add hover effect for value columns only
            Trigger mouseOverTrigger = new Trigger { Property = DataGridCell.IsMouseOverProperty, Value = true };
            mouseOverTrigger.Setters.Add(new Setter(DataGridCell.BackgroundProperty, Brushes.LightBlue));
            valueColumnStyle.Triggers.Add(mouseOverTrigger);

            // Apply column-specific styles
            _valueDataGrid.Columns[0].CellStyle = concreteTypeColumnStyle; // Concrete Type column
            _valueDataGrid.Columns[1].CellStyle = valueColumnStyle; // 0% column
            _valueDataGrid.Columns[2].CellStyle = valueColumnStyle; // 25% column  
            _valueDataGrid.Columns[3].CellStyle = valueColumnStyle; // 50% column
            _valueDataGrid.Columns[4].CellStyle = valueColumnStyle; // 75% column


            // Add event handler for DataGrid
            _valueDataGrid.MouseLeftButtonUp += ValueDataGrid_MouseLeftButtonUp;
            _valueDataGrid.PreviewKeyDown += ValueDataGrid_PreviewKeyDown;

            // Create Grid to hold DataGrid and make it fill the space
            var grid = (Grid)border.Child;
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            // Add DataGrid to the Grid and make it fill the space
            Grid.SetRow(_valueDataGrid, 0);
            grid.Children.Add(_valueDataGrid);

            // Add title/header
            TextBlock headerText = new TextBlock
            {
                Text = "Concrete Embodied Carbon Values (kgCO2e/m³)",
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
        }

        public void SetValues(IEnumerable<ConcreteData> values)
        {
            _valueDataGrid.ItemsSource = values;
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
            // Handle Enter key to select the current cell
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
            ConcreteData item = cellInfo.Item as ConcreteData;

            // Only allow selecting cells from percentage columns, not from the concrete type column
            if (item != null && headerName != "Concrete Type")
            {
                string value = "";

                if (headerName == "0%")
                    value = item.ZeroPercent;
                else if (headerName == "25%")
                    value = item.TwentyFivePercent;
                else if (headerName == "50%")
                    value = item.FiftyPercent;
                else if (headerName == "75%")
                    value = item.SeventyFivePercent;

                if (!string.IsNullOrEmpty(value))
                {
                    ValueSelected?.Invoke(this, value);
                }
            }
            // If the concrete type column was clicked, don't do anything (don't trigger ValueSelected)
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
                // Skip if this is the concrete type column (column index 0)
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

                    // Skip the concrete type column (column index 0)
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

    public class ConcreteData
    {
        public string ConcreteType { get; set; }
        public string ZeroPercent { get; set; }
        public string TwentyFivePercent { get; set; }
        public string FiftyPercent { get; set; }
        public string SeventyFivePercent { get; set; }
    }


}