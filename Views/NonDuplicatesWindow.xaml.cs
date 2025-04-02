using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Microsoft.Win32;

namespace RffaDataComparisonTool.Views
{
    public partial class NonDuplicatesWindow : Window
    {
        private Dictionary<string, List<string>> _nonDuplicatesBySheet;
        private int _totalNonDuplicates;

        public NonDuplicatesWindow()
        {
            InitializeComponent();

            // Prevent window from minimizing main window on close
            this.Closing += (s, e) => this.Owner = null;
        }

        public int TotalNonDuplicates
        {
            get { return _totalNonDuplicates; }
        }

        /// <summary>
        /// Initialize the window with data safely
        /// </summary>
        public void Initialize(Dictionary<string, List<string>> nonDuplicatesBySheet)
        {
            try
            {
                _nonDuplicatesBySheet = nonDuplicatesBySheet ?? new Dictionary<string, List<string>>();
                _totalNonDuplicates = _nonDuplicatesBySheet.Values.Sum(list => list.Count);

                // Set DataContext to this for simple binding
                DataContext = this;

                // Populate the content
                PopulateContent();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing non-duplicates window: {ex.Message}",
                                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PopulateContent()
        {
            try
            {
                // Clear existing content
                ContentPanel.Children.Clear();

                // Add each sheet section
                foreach (var sheet in _nonDuplicatesBySheet)
                {
                    // Create sheet header
                    TextBlock header = new TextBlock
                    {
                        Text = $"Sheet: {sheet.Key}",
                        FontWeight = FontWeights.Bold,
                        Foreground = new SolidColorBrush(Color.FromRgb(25, 118, 210)), // #1976D2
                        Margin = new Thickness(0, 10, 0, 5)
                    };
                    ContentPanel.Children.Add(header);

                    // Add each RSBSA number
                    foreach (var value in sheet.Value)
                    {
                        TextBlock item = new TextBlock
                        {
                            Text = value,
                            Margin = new Thickness(15, 2, 0, 2)
                        };
                        ContentPanel.Children.Add(item);
                    }

                    // Add separator
                    if (sheet.Key != _nonDuplicatesBySheet.Keys.Last())
                    {
                        Separator separator = new Separator
                        {
                            Margin = new Thickness(0, 10, 0, 10)
                        };
                        ContentPanel.Children.Add(separator);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error populating content: {ex.Message}",
                                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_nonDuplicatesBySheet == null || _totalNonDuplicates == 0)
                {
                    MessageBox.Show("No data to copy.", "Information",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Build the text to copy
                string text = string.Empty;
                foreach (var sheet in _nonDuplicatesBySheet)
                {
                    foreach (var value in sheet.Value)
                    {
                        text += value + Environment.NewLine;
                    }
                }

                // Copy to clipboard
                Clipboard.SetText(text);
                MessageBox.Show("Copied to clipboard successfully!", "Success",
                               MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying to clipboard: {ex.Message}",
                               "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_nonDuplicatesBySheet == null || _totalNonDuplicates == 0)
                {
                    MessageBox.Show("No data to export.", "Information",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Create save file dialog
                SaveFileDialog dialog = new SaveFileDialog
                {
                    Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                    DefaultExt = ".txt",
                    FileName = "Non_Duplicate_RSBSA_Numbers.txt"
                };

                if (dialog.ShowDialog() == true)
                {
                    using (StreamWriter writer = new StreamWriter(dialog.FileName))
                    {
                        writer.WriteLine("Non-duplicate RSBSA Reference Numbers:");
                        writer.WriteLine();

                        foreach (var sheet in _nonDuplicatesBySheet)
                        {
                            writer.WriteLine($"Sheet: {sheet.Key}");
                            foreach (var value in sheet.Value)
                            {
                                writer.WriteLine(value);
                            }
                            writer.WriteLine();  // Empty line between sheets
                        }
                    }

                    MessageBox.Show($"File exported successfully to:\n{dialog.FileName}",
                                  "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting file: {ex.Message}",
                               "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}