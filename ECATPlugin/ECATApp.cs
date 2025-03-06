using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

namespace ECATPlugin
{
    public class ECATApp : IExternalApplication
    {
        // Static reference to maintain a single instance of the window
        private static MainWindow _mainWindow = null;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create a Ribbon Panel in the Revit UI
                RibbonPanel ribbonPanel = application.CreateRibbonPanel("Walsh ECAT");

                // Get the path of this assembly
                string thisAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

                // Define the data for the push button
                PushButtonData buttonData = new PushButtonData(
                    "cmdWalshECAT",          // Unique button identifier
                    "Walsh ECAT",            // Button text
                    thisAssemblyPath,
                    "ECATPlugin.WalshECAT"   // Full class name for the command to execute
                );

                // Add the push button to the ribbon panel
                PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;

                if (pushButton != null)
                {
                    pushButton.ToolTip = "Launch the Walsh ECAT Tool"; // Add a tooltip for clarity
                }

                pushButton.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(@"C:\Users\hummd\source\repos\ECATPlugin\ECATPlugin\bin\Debug\walsh_logo.png"));

                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to initialize Walsh ECAT Plugin: {ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // Close the window if it's open
            if (_mainWindow != null)
            {
                _mainWindow.Close();
                _mainWindow = null;
            }

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class WalshECAT : IExternalCommand
    {
        // Static reference to keep track of the open window
        private static MainWindow _mainWindow = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiApp = commandData.Application;
            var doc = uiApp.ActiveUIDocument.Document;

            try
            {
                // If the window is already open, just activate it
                if (_mainWindow != null)
                {
                    _mainWindow.Activate();
                    return Result.Succeeded;
                }

                // Create an instance of MainViewModel and pass the Revit Document
                MainViewModel viewModel = new MainViewModel(doc);

                // Instantiate the WPF Window
                _mainWindow = new MainWindow(viewModel);

                // Set the DataContext of the WPF Window to the ViewModel
                _mainWindow.DataContext = viewModel;

                // Set up event handler to nullify the reference when the window is closed
                _mainWindow.Closed += (s, e) =>
                {
                    _mainWindow = null;
                };

                // Show the WPF Window as a modeless dialog
                _mainWindow.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}