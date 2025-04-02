using RffaDataComparisonTool.Helpers;
using RffaDataComparisonTool.Models;
using RffaDataComparisonTool.Services;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace RffaDataComparisonTool.ViewModels
{
    public class InvalidFarmAreaViewModel : ObservableObject
    {
        private readonly List<FarmAreaRecord> _records;
        private readonly string _rffaFilePath;
        private readonly ExcelProcessorService _excelProcessor;
        private int _currentIndex = 0;

        // Properties
        private FarmAreaRecord _currentRecord;
        public FarmAreaRecord CurrentRecord
        {
            get => _currentRecord;
            private set => SetProperty(ref _currentRecord, value, nameof(CurrentRecord));
        }

        // Display properties
        public string CurrentRecordText => _records.Count > 0 ?
            $"{_currentIndex + 1} of {_records.Count}" : "No records";

        public string ProgressText => _records.Count > 0 ?
            $"Viewing record {_currentIndex + 1} of {_records.Count}" : "";

        public string IssueDescription
        {
            get
            {
                if (CurrentRecord == null) return string.Empty;

                if (CurrentRecord.FarmArea > 2.0)
                    return $"Farm area exceeds maximum allowed value (2.0 Ha).";
                else if (CurrentRecord.FarmArea < 0)
                    return $"Farm area cannot be negative.";
                else if (CurrentRecord.FarmArea == 0)
                    return $"Farm area cannot be zero.";

                return "Valid farm area.";
            }
        }

        public Brush FarmAreaColor
        {
            get
            {
                if (CurrentRecord == null) return Brushes.Black;

                if (CurrentRecord.FarmArea > 2.0)
                    return Brushes.DarkOrange;
                else if (CurrentRecord.FarmArea < 0)
                    return Brushes.Red;
                else if (CurrentRecord.FarmArea == 0)
                    return Brushes.Red;

                return Brushes.Green;
            }
        }

        // Visibility properties
        public Visibility NoRecordsVisibility => _records.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        public Visibility RecordDetailsVisibility => _records.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        public Visibility NavigationButtonsVisibility => _records.Count > 1 ? Visibility.Visible : Visibility.Collapsed;
        public Visibility ProgressVisibility => _records.Count > 0 ? Visibility.Visible : Visibility.Collapsed;

        // Commands
        public ICommand AcceptAndIgnoreCommand { get; }
        public ICommand RemoveHighlightCommand { get; }
        public ICommand ExportToExcelCommand { get; }
        public ICommand CopyToClipboardCommand { get; }
        public ICommand CloseCommand { get; }
        public ICommand NextRecordCommand { get; }
        public ICommand PreviousRecordCommand { get; }

        // Dialog result for window closure
        private bool? _dialogResult;
        public bool? DialogResult
        {
            get => _dialogResult;
            set => SetProperty(ref _dialogResult, value, nameof(DialogResult));
        }

        public InvalidFarmAreaViewModel(List<FarmAreaRecord> records, string rffaFilePath)
        {
            _records = records ?? new List<FarmAreaRecord>();
            _rffaFilePath = rffaFilePath;
            _excelProcessor = new ExcelProcessorService(
                new HistoryService(),
                new UserPreferencesService()
            );

            // Set initial record if available
            if (_records.Count > 0)
            {
                CurrentRecord = _records[0];
            }

            // Initialize commands
            AcceptAndIgnoreCommand = new RelayCommand(_ => AcceptAndIgnore());
            RemoveHighlightCommand = new RelayCommand(_ => RemoveHighlight());
            ExportToExcelCommand = new RelayCommand(_ => ExportToExcel(), _ => _records.Count > 0);
            CopyToClipboardCommand = new RelayCommand(_ => CopyToClipboard(), _ => _records.Count > 0);
            CloseCommand = new RelayCommand(_ => Close());
            NextRecordCommand = new RelayCommand(_ => NextRecord(), _ => CanGoToNextRecord());
            PreviousRecordCommand = new RelayCommand(_ => PreviousRecord(), _ => CanGoToPreviousRecord());
        }

        private void AcceptAndIgnore()
        {
            if (CurrentRecord == null) return;

            // Ask for confirmation
            var result = MessageBox.Show(
                "Are you sure you want to accept this record and ignore the error? The row will remain highlighted in the Excel file.",
                "Confirm Ignore Error",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return;

            try
            {
                // Mark record as not highlighted (in our tracking, but don't remove from Excel)
                CurrentRecord.IsHighlighted = false;

                // Go to next record automatically
                if (CanGoToNextRecord())
                {
                    // Use Application.Current.Dispatcher to avoid Owner issue
                    Application.Current.Dispatcher.BeginInvoke(new Action(() => {
                        NextRecord();
                    }));
                }
                else
                {
                    // Use Application.Current.Dispatcher to avoid Owner issue
                    Application.Current.Dispatcher.BeginInvoke(new Action(() => {
                        MessageBox.Show("All records processed.", "Complete",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                        Close();
                    }));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void RemoveHighlight()
        {
            if (CurrentRecord == null) return;

            try
            {
                // Remove highlight from the Excel file
                bool success = await _excelProcessor.RemoveHighlightAsync(_rffaFilePath, CurrentRecord);

                if (success)
                {
                    // Mark record as not highlighted
                    CurrentRecord.IsHighlighted = false;

                    MessageBox.Show("Highlight removed successfully from the Excel file.",
                        "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                    // Go to next record automatically
                    if (CanGoToNextRecord())
                    {
                        // Use Application.Current.Dispatcher to avoid Owner issue
                        Application.Current.Dispatcher.BeginInvoke(new Action(() => {
                            NextRecord();
                        }));
                    }
                    else
                    {
                        // Use Application.Current.Dispatcher to avoid Owner issue
                        Application.Current.Dispatcher.BeginInvoke(new Action(() => {
                            MessageBox.Show("All records processed.", "Complete",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                            Close();
                        }));
                    }
                }
                else
                {
                    MessageBox.Show("Failed to update the Excel file. Please make sure it is not open in another application.",
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void ExportToExcel()
        {
            try
            {
                // Use save dialog to get the export location
                var dialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    DefaultExt = ".xlsx",
                    FileName = "Invalid_Farm_Areas.xlsx"
                };

                if (dialog.ShowDialog() == true)
                {
                    string saveLocation = System.IO.Path.GetDirectoryName(dialog.FileName);
                    string exportPath = await _excelProcessor.ExportInvalidFarmAreasToExcel(_records, saveLocation);

                    MessageBox.Show($"Records exported to:\n{exportPath}",
                        "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CopyToClipboard()
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                // Add header
                sb.AppendLine("Sheet\tRSBSA Number\tLast Name\tFirst Name\tMiddle Name\tFarm Area (Ha)");

                // Add data
                foreach (var record in _records)
                {
                    sb.AppendLine($"{record.SheetName}\t{record.RsbsaNumber}\t{record.LastName}\t{record.FirstName}\t{record.MiddleName}\t{record.FarmArea}");
                }

                // Copy to clipboard
                Clipboard.SetText(sb.ToString());

                MessageBox.Show("Data copied to clipboard.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to copy to clipboard: {ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void NextRecord()
        {
            if (CanGoToNextRecord())
            {
                _currentIndex++;
                CurrentRecord = _records[_currentIndex];
                UpdateProperties();
            }
        }

        private void PreviousRecord()
        {
            if (CanGoToPreviousRecord())
            {
                _currentIndex--;
                CurrentRecord = _records[_currentIndex];
                UpdateProperties();
            }
        }

        private bool CanGoToNextRecord()
        {
            return _records.Count > 0 && _currentIndex < _records.Count - 1;
        }

        private bool CanGoToPreviousRecord()
        {
            return _records.Count > 0 && _currentIndex > 0;
        }

        private void Close()
        {
            DialogResult = false;
        }

        private void UpdateProperties()
        {
            // Notify property changes that depend on CurrentRecord
            OnPropertyChanged(nameof(CurrentRecordText));
            OnPropertyChanged(nameof(ProgressText));
            OnPropertyChanged(nameof(IssueDescription));
            OnPropertyChanged(nameof(FarmAreaColor));

            // Invalidate command can-execute state
            CommandManager.InvalidateRequerySuggested();
        }
    }
}