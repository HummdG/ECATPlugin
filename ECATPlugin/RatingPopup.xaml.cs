using System;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using Autodesk.Revit.DB;

namespace ECATPlugin
{
    public partial class RatingPopup : Window, INotifyPropertyChanged
    {
        private double _carbonRating;
        public double CarbonRating
        {
            get => _carbonRating;
            set
            {
                _carbonRating = value;
                UpdateArrowPosition(RatingType.Structe); // Update Structe arrow position
                UpdateArrowPosition(RatingType.Leti);    // Update Leti arrow position
            }
        }

        private string _structeGrade;
        public string StructeGrade
        {
            get => _structeGrade;
            set
            {
                _structeGrade = value;
                OnPropertyChanged(nameof(StructeGrade));
            }
        }

        private string _letiGrade;
        public string LetiGrade
        {
            get => _letiGrade;
            set
            {
                _letiGrade = value;
                OnPropertyChanged(nameof(LetiGrade));
            }
        }
        private MainViewModel _viewModel;
        public RatingPopup(double carbonRating, MainViewModel viewModel)
        {
            InitializeComponent();
            DataContext = this;
            CarbonRating = carbonRating;
            _viewModel = viewModel;
        }

        private void UpdateArrowPosition(RatingType ratingType)
        {
            double topPosition = CalculateTopPosition(CarbonRating, ratingType);
            PositionArrow(topPosition, ratingType);
            string grade = CalculateGrade(CarbonRating, ratingType);

            if (ratingType == RatingType.Structe)
                StructeGrade = grade;
            else
                LetiGrade = grade;
        }

        private double CalculateTopPosition(double rating, RatingType ratingType)
        {
            if (ratingType == RatingType.Structe)
            {
                if (rating <= 100) return 12.5;
                if (rating <= 150) return 42.5;
                if (rating <= 200) return 72.5;
                if (rating <= 250) return 102.5;
                if (rating <= 300) return 132.5;
                if (rating <= 400) return 162.5;
                if (rating <= 500) return 192.5;
                if (rating <= 625) return 222.5;
                return 252.5;
            }
            else // Leti
            {
                if (rating <= 100) return 12.5;
                if (rating <= 200) return 42.5;
                if (rating <= 350) return 72.5;
                if (rating <= 500) return 102.5;
                if (rating <= 800) return 132.5;
                if (rating <= 1000) return 162.5;
                if (rating <= 1200) return 192.5;
                if (rating <= 1400) return 222.5;
                return 252.5;
            }
        }

        private string CalculateGrade(double rating, RatingType ratingType)
        {
            if (ratingType == RatingType.Structe)
            {
                if (rating <= 100) return "A++";
                if (rating <= 150) return "A+";
                if (rating <= 200) return "A";
                if (rating <= 250) return "B";
                if (rating <= 300) return "C";
                if (rating <= 400) return "D";
                if (rating <= 500) return "E";
                if (rating <= 625) return "F";
                return "G";
            }
            else // Leti
            {
                if (rating <= 100) return "A++";
                if (rating <= 200) return "A+";
                if (rating <= 350) return "A";
                if (rating <= 500) return "B";
                if (rating <= 800) return "C";
                if (rating <= 1000) return "D";
                if (rating <= 1200) return "E";
                if (rating <= 1400) return "F";
                return "G";
            }
        }

        private void PositionArrow(double topPosition, RatingType ratingType)
        {
            if (ratingType == RatingType.Structe)
            {
                Canvas.SetLeft(StructeArrowPolygon, 280);
                Canvas.SetTop(StructeArrowPolygon, topPosition);
                Canvas.SetLeft(StructeRatingLabelBorder, 298);
                Canvas.SetTop(StructeRatingLabelBorder, topPosition + 1);
            }
            else if (ratingType == RatingType.Leti)
            {
                Canvas.SetLeft(LetiArrowPolygon, 227);
                Canvas.SetTop(LetiArrowPolygon, topPosition);
                Canvas.SetLeft(LetiRatingLabelBorder, 245);
                Canvas.SetTop(LetiRatingLabelBorder, topPosition + 1);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private enum RatingType
        {
            Structe,
            Leti
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
                string fileName = $"Ratings_{_viewModel.ProjectNumber}_{_viewModel.ProjectName}.png";

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
    }
}
