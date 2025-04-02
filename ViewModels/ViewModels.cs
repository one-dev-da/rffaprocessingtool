using Microsoft.Win32;
using RffaDataComparisonTool.Helpers;
using RffaDataComparisonTool.Models;
using RffaDataComparisonTool.Services;
using RffaDataComparisonTool.Views;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace RffaDataComparisonTool.ViewModels
{
    /// <summary>
    /// ViewModel for the main window
    /// </summary>
    public class MainViewModel : ObservableObject
    {
        private readonly ExcelProcessorService _excelProcessor;
        private readonly HistoryService _historyService;
        private readonly ILoadingService _loadingService;
        private readonly UserPreferencesService _preferencesService;

        // Selected history items for deletion
        private List<ProcessingRecord> _selectedHistoryItems = new List<ProcessingRecord>();
        public bool BothFilesSelected => !string.IsNullOrEmpty(MagaraoPath) && !string.IsNullOrEmpty(ImpTopupPath);

        // File paths
        private string _magaraoPath;
        public string MagaraoPath
        {
            get => _magaraoPath;
            set
            {
                if (SetProperty(ref _magaraoPath, value, nameof(MagaraoPath)))
                {
                    OnPropertyChanged(nameof(MagaraoFileName));
                    OnPropertyChanged(nameof(IsMagaraoFileSelected));
                    OnPropertyChanged(nameof(BothFilesSelected));
                    OnPropertyChanged(nameof(CanProcess));

                    // Only load sheets if both files are selected
                    if (!string.IsNullOrEmpty(ImpTopupPath))
                    {
                        LoadSheets();
                    }
                }
            }
        }

        private string _impTopupPath;
        public string ImpTopupPath
        {
            get => _impTopupPath;
            set
            {
                if (SetProperty(ref _impTopupPath, value, nameof(ImpTopupPath)))
                {
                    OnPropertyChanged(nameof(ImpTopupFileName));
                    OnPropertyChanged(nameof(IsImpTopupFileSelected));
                    OnPropertyChanged(nameof(BothFilesSelected));
                    OnPropertyChanged(nameof(CanProcess));

                    // If RFFA file is already selected, load sheets now that both files are selected
                    if (!string.IsNullOrEmpty(MagaraoPath))
                    {
                        LoadSheets();
                    }
                }
            }
        }

        // File names (for display)
        public string MagaraoFileName => !string.IsNullOrEmpty(MagaraoPath) ?
            Path.GetFileName(MagaraoPath) : "Not selected";

        public string ImpTopupFileName => !string.IsNullOrEmpty(ImpTopupPath) ?
            Path.GetFileName(ImpTopupPath) : "Not selected";

        // Sheet selection
        private ObservableCollection<SheetInfo> _availableSheets = new ObservableCollection<SheetInfo>();
        public ObservableCollection<SheetInfo> AvailableSheets => _availableSheets;

        private SheetInfo _selectedSheet;
        public SheetInfo SelectedSheet
        {
            get => _selectedSheet;
            set
            {
                if (SetProperty(ref _selectedSheet, value, nameof(SelectedSheet)))
                {
                    OnPropertyChanged(nameof(CanProcess));
                }
            }
        }

        private ObservableCollection<string> _selectedSheets = new ObservableCollection<string>();
        public ObservableCollection<string> SelectedSheets => _selectedSheets;

        // Processing results
        private int _duplicatesFound;
        public int DuplicatesFound
        {
            get => _duplicatesFound;
            set
            {
                if (SetProperty(ref _duplicatesFound, value, nameof(DuplicatesFound)))
                {
                    // When DuplicatesFound changes, Total Endorsed also changes
                    OnPropertyChanged(nameof(TotalEndorsed));
                }
            }
        }

        private int _nonDuplicatesFound;
        public int NonDuplicatesFound
        {
            get => _nonDuplicatesFound;
            set => SetProperty(ref _nonDuplicatesFound, value, nameof(NonDuplicatesFound));
        }

        private int _invalidFarmAreasFound;
        public int InvalidFarmAreasFound
        {
            get => _invalidFarmAreasFound;
            set
            {
                if (SetProperty(ref _invalidFarmAreasFound, value, nameof(InvalidFarmAreasFound)))
                {
                    // When InvalidFarmAreasFound changes, Total Endorsed also changes
                    OnPropertyChanged(nameof(TotalEndorsed));
                }
            }
        }

        private int _totalEndorsed;
        public int TotalEndorsed
        {
            get => _totalEndorsed;
            private set => SetProperty(ref _totalEndorsed, value, nameof(TotalEndorsed));
        }       

        // Total Endorsed is Duplicates minus Invalid Farm Areas
        //public int TotalEndorsed => DuplicatesFound - InvalidFarmAreasFound;

        // Total Highlighted RSBSA in RFFA
        private int _totalHighlightedRsbsa;
        public int TotalHighlightedRsbsa
        {
            get => _totalHighlightedRsbsa;
            set => SetProperty(ref _totalHighlightedRsbsa, value, nameof(TotalHighlightedRsbsa));
        }

        private string _saveLocation;
        public string SaveLocation
        {
            get => _saveLocation;
            set => SetProperty(ref _saveLocation, value, nameof(SaveLocation));
        }

        private ObservableCollection<string> _duplicatesList = new ObservableCollection<string>();
        public ObservableCollection<string> DuplicatesList => _duplicatesList;

        private Dictionary<string, List<string>> _nonDuplicatesBySheet = new Dictionary<string, List<string>>();
        public Dictionary<string, List<string>> NonDuplicatesBySheet => _nonDuplicatesBySheet;

        private List<FarmAreaRecord> _invalidFarmAreaRecords = new List<FarmAreaRecord>();
        public List<FarmAreaRecord> InvalidFarmAreaRecords => _invalidFarmAreaRecords;

        // Property to check if Magarao file is selected
        public bool IsMagaraoFileSelected => !string.IsNullOrEmpty(MagaraoPath);

        // Property to check if IMP file is selected
        public bool IsImpTopupFileSelected => !string.IsNullOrEmpty(ImpTopupPath);

        private int _currentPage = 1;
        private int _pageSize = 100;
        private ObservableCollection<string> _currentPageItems = new ObservableCollection<string>();

        public ObservableCollection<string> CurrentPageItems => _currentPageItems;

        public string PaginationInfo => $"Page {_currentPage} of {TotalPages} ({DuplicatesList.Count} items)";

        public int TotalPages => (DuplicatesList.Count + _pageSize - 1) / _pageSize; // Ceiling division

        // Status
        private string _statusMessage = "Select all required files to begin";
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value, nameof(StatusMessage));
        }

        // History
        public ObservableCollection<ProcessingRecord> History => _historyService.History;

        // Button states
        public bool CanProcess => !string.IsNullOrEmpty(MagaraoPath) &&
                          !string.IsNullOrEmpty(ImpTopupPath) &&
                          (SelectedSheet != null || SelectedSheets.Count > 0);

        private bool _canOpenImpTopupFile;
        public bool CanOpenImpTopupFile
        {
            get => _canOpenImpTopupFile;
            set => SetProperty(ref _canOpenImpTopupFile, value, nameof(CanOpenImpTopupFile));
        }

        private bool _canOpenRffaFile;
        public bool CanOpenRffaFile
        {
            get => _canOpenRffaFile;
            set => SetProperty(ref _canOpenRffaFile, value, nameof(CanOpenRffaFile));
        }

        private bool _canViewNonDuplicates;
        public bool CanViewNonDuplicates
        {
            get => _canViewNonDuplicates;
            set => SetProperty(ref _canViewNonDuplicates, value, nameof(CanViewNonDuplicates));
        }

        private bool _canViewInvalidFarmAreas;
        public bool CanViewInvalidFarmAreas
        {
            get => _canViewInvalidFarmAreas;
            set => SetProperty(ref _canViewInvalidFarmAreas, value, nameof(CanViewInvalidFarmAreas));
        }

        private bool _canExportBatch;
        public bool CanExportBatch
        {
            get => _canExportBatch;
            set => SetProperty(ref _canExportBatch, value, nameof(CanExportBatch));
        }

        // Commands
        public ICommand BrowseMagaraoCommand { get; }
        public ICommand BrowseImpTopupCommand { get; }
        public ICommand ProcessFilesCommand { get; }
        public ICommand ExportHistoryCommand { get; }
        public ICommand OpenImpTopupFileCommand { get; }
        public ICommand OpenRffaFileCommand { get; }
        public ICommand ViewNonDuplicatesCommand { get; }
        public ICommand ViewInvalidFarmAreasCommand { get; }
        public ICommand DeleteSelectedHistoryCommand { get; }
        public ICommand ClearAllHistoryCommand { get; }
        public ICommand SelectMultipleSheetsCommand { get; }
        public ICommand CopyDuplicatesCommand { get; }
        public ICommand FirstPageCommand { get; }
        public ICommand PreviousPageCommand { get; }
        public ICommand NextPageCommand { get; }
        public ICommand LastPageCommand { get; }
        public ICommand ExportBatchCommand { get; }

        public MainViewModel(ExcelProcessorService excelProcessor, HistoryService historyService, ILoadingService loadingService, UserPreferencesService preferencesService)
        {
            _excelProcessor = excelProcessor;
            _historyService = historyService;
            _loadingService = loadingService;
            _preferencesService = preferencesService;

            // Initialize commands constructor
            BrowseMagaraoCommand = new RelayCommand(_ => BrowseMagaraoFile());
            BrowseImpTopupCommand = new RelayCommand(_ => BrowseImpTopupFile());
            ProcessFilesCommand = new RelayCommand(_ => ProcessFiles(), _ => CanProcess);
            ExportHistoryCommand = new RelayCommand(_ => ExportHistory(), _ => History.Count > 0);
            OpenImpTopupFileCommand = new RelayCommand(_ => OpenImpTopupFile(), _ => CanOpenImpTopupFile);
            OpenRffaFileCommand = new RelayCommand(_ => OpenRffaFile(), _ => CanOpenRffaFile);
            ViewNonDuplicatesCommand = new RelayCommand(_ => ViewNonDuplicates(), _ => CanViewNonDuplicates);
            ViewInvalidFarmAreasCommand = new RelayCommand(_ => ViewInvalidFarmAreas(), _ => CanViewInvalidFarmAreas);
            DeleteSelectedHistoryCommand = new RelayCommand(_ => DeleteSelectedHistory(), _ => _selectedHistoryItems.Count > 0);
            ClearAllHistoryCommand = new RelayCommand(_ => ClearAllHistory(), _ => History.Count > 0);
            SelectMultipleSheetsCommand = new RelayCommand(_ => SelectMultipleSheets(), _ => AvailableSheets.Count > 0);
            CopyDuplicatesCommand = new RelayCommand(_ => CopyDuplicatesToClipboard(), _ => DuplicatesList.Count > 0);
            FirstPageCommand = new RelayCommand(_ => GoToFirstPage(), _ => CanGoToFirstPage());
            PreviousPageCommand = new RelayCommand(_ => GoToPreviousPage(), _ => CanGoToPreviousPage());
            NextPageCommand = new RelayCommand(_ => GoToNextPage(), _ => CanGoToNextPage());
            LastPageCommand = new RelayCommand(_ => GoToLastPage(), _ => CanGoToLastPage());
            ExportBatchCommand = new RelayCommand(_ => ExportBatch(), _ => CanExportBatch);
        }

        private void ExportBatch()
        {
            if (string.IsNullOrEmpty(MagaraoPath) || !File.Exists(MagaraoPath))
            {
                MessageBox.Show("RFFA file not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrEmpty(ImpTopupPath) || !File.Exists(ImpTopupPath))
            {
                MessageBox.Show("IMP Top-up file not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // Create a batch export dialog, passing both file paths
                var batchWindow = new Views.BatchExportWindow(MagaraoPath, ImpTopupPath, SaveLocation);

                // Show the window
                bool? result = batchWindow.ShowDialog();

                if (result == true)
                {
                    // Batch export successful
                    string batchFilePath = batchWindow.GeneratedBatchFilePath;

                    var messageResult = MessageBox.Show(
                        $"Batch file exported successfully!\n\nWould you like to open the file now?",
                        "Export Complete",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (messageResult == MessageBoxResult.Yes)
                    {
                        try
                        {
                            Process.Start(new ProcessStartInfo(batchFilePath) { UseShellExecute = true });
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Failed to open file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening batch export window: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CopyDuplicatesToClipboard()
        {
            if (DuplicatesList.Count == 0)
            {
                MessageBox.Show("No duplicate RSBSA numbers to copy.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                // Create a string with all duplicates, removing any numbering/formatting
                var duplicatesText = string.Join(Environment.NewLine,
                    DuplicatesList
                        .Select(item => item.Contains(". ")
                            ? item.Substring(item.IndexOf(". ") + 2)  // Remove numbering (e.g., "1. ")
                            : item));

                // Copy to clipboard
                Clipboard.SetText(duplicatesText);
                MessageBox.Show("Duplicate RSBSA numbers copied to clipboard!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to copy to clipboard: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateCurrentPageItems()
        {
            _currentPageItems.Clear();

            if (DuplicatesList.Count == 0)
                return;

            int startIndex = (_currentPage - 1) * _pageSize;
            int itemsToTake = Math.Min(_pageSize, DuplicatesList.Count - startIndex);

            if (startIndex < DuplicatesList.Count)
            {
                for (int i = startIndex; i < startIndex + itemsToTake; i++)
                {
                    _currentPageItems.Add(DuplicatesList[i]);
                }
            }

            OnPropertyChanged(nameof(PaginationInfo));
            OnPropertyChanged(nameof(CurrentPageItems));

            // Make sure command availability is updated
            CommandManager.InvalidateRequerySuggested();
        }


        private void GoToFirstPage()
        {
            _currentPage = 1;
            UpdateCurrentPageItems();
        }

        private bool CanGoToFirstPage() => _currentPage > 1;

        private void GoToPreviousPage()
        {
            if (_currentPage > 1)
            {
                _currentPage--;
                UpdateCurrentPageItems();
            }
        }

        private bool CanGoToPreviousPage() => _currentPage > 1;

        private void GoToNextPage()
        {
            if (_currentPage < TotalPages)
            {
                _currentPage++;
                UpdateCurrentPageItems();
            }
        }

        private bool CanGoToNextPage() => _currentPage < TotalPages && DuplicatesList.Count > 0;

        private void GoToLastPage()
        {
            _currentPage = TotalPages;
            UpdateCurrentPageItems();
        }

        private bool CanGoToLastPage() => _currentPage < TotalPages && DuplicatesList.Count > 0;

        /// <summary>
        /// Updates the selected history items and refreshes the delete command
        /// </summary>
        public void UpdateSelectedHistoryItems(List<ProcessingRecord> items)
        {
            _selectedHistoryItems = items;

            // Re-evaluate CanExecute for DeleteSelectedHistoryCommand
            CommandManager.InvalidateRequerySuggested();
        }

        /// <summary>
        /// Checks if a file is locked/in use by another process
        /// </summary>
        /// <param name="filePath">Path to the file to check</param>
        /// <returns>True if the file is locked, false otherwise</returns>
        private bool IsFileLocked(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return false;

            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // If we get here, the file is not locked
                    stream.Close();
                    return false;
                }
            }
            catch (IOException)
            {
                // The file is locked by another process
                return true;
            }
            catch
            {
                // Another exception occurred, assume not locked
                return false;
            }
        }

        private void BrowseMagaraoFile()
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select RFFA file",
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                CheckFileExists = true
            };

            if (dialog.ShowDialog() == true)
            {
                MagaraoPath = dialog.FileName;
                StatusMessage = "RFFA file selected.";

                if (!string.IsNullOrEmpty(ImpTopupPath))
                {
                    StatusMessage = "All files selected. Ready to process.";
                }
                else
                {
                    StatusMessage = "Please select IMP Topup Module file.";
                }
            }
        }

        private void BrowseImpTopupFile()
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select IMP Topup Module file",
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                CheckFileExists = true
            };

            if (dialog.ShowDialog() == true)
            {
                ImpTopupPath = dialog.FileName;

                if (!string.IsNullOrEmpty(MagaraoPath))
                {
                    StatusMessage = "All files selected. Ready to process.";
                }
                else
                {
                    StatusMessage = "Please select RFFA file.";
                }
            }
        }

        private void LoadSheets()
        {
            if (string.IsNullOrEmpty(MagaraoPath) || !File.Exists(MagaraoPath) || string.IsNullOrEmpty(ImpTopupPath))
                return;

            try
            {
                var sheets = _excelProcessor.GetVisibleSheets(MagaraoPath);

                _availableSheets.Clear();
                foreach (var sheet in sheets)
                {
                    _availableSheets.Add(sheet);
                }

                if (_availableSheets.Count > 0)
                {
                    SelectedSheet = _availableSheets[0];
                }

                OnPropertyChanged(nameof(AvailableSheets));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load sheets: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SelectMultipleSheets()
        {
            try
            {
                var viewModel = new SheetSelectionViewModel(AvailableSheets.ToList());
                var window = new Views.SheetSelectionWindow
                {
                    DataContext = viewModel,
                    Owner = Application.Current.MainWindow
                };

                if (window.ShowDialog() == true && viewModel.SelectedSheets.Count > 0)
                {
                    _selectedSheets.Clear();
                    foreach (var sheet in viewModel.SelectedSheets)
                    {
                        _selectedSheets.Add(sheet);
                    }

                    StatusMessage = $"Selected {_selectedSheets.Count} sheet(s) for processing.";
                    OnPropertyChanged(nameof(CanProcess));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error selecting sheets: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void ProcessFiles()
        {
            try
            {
                // Show loading overlay using the service
                _loadingService.ShowLoading("Processing files...");

                if (IsFileLocked(ImpTopupPath))
                {
                    MessageBox.Show("The IMP Topup file is currently open.\nPlease close the file before processing.",
                        "File In Use", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _loadingService.HideLoading();
                    return;
                }

                if (IsFileLocked(MagaraoPath))
                {
                    MessageBox.Show("The RFFA file is currently open.\nPlease close the file before processing.",
                        "File In Use", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _loadingService.HideLoading();
                    return;
                }

                List<string> sheetsToProcess;

                // Use selected sheets from multi-select if available, otherwise use the single selection
                if (SelectedSheets.Count > 0)
                {
                    sheetsToProcess = SelectedSheets.ToList();
                }
                else if (SelectedSheet != null)
                {
                    sheetsToProcess = new List<string> { SelectedSheet.Name };
                }
                else
                {
                    MessageBox.Show("Please select at least one sheet to process.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    _loadingService.HideLoading();
                    return;
                }

                // Create a progress reporter for the log window
                var logWindow = new Views.ProcessingLogWindow();
                logWindow.Owner = Application.Current.MainWindow;
                logWindow.Show();

                var progress = new Progress<string>(message =>
                {
                    logWindow.AppendLog(message);
                });

                StatusMessage = "Processing files... Please wait.";

                // Ask about backup preference if needed
                bool createBackup = _preferencesService.AlwaysCreateBackup;

                if (!_preferencesService.AlwaysCreateBackup)
                {
                    var backupWindow = new Views.BackupConfirmWindow();
                    backupWindow.Owner = Application.Current.MainWindow;

                    if (backupWindow.ShowDialog() == true)
                    {
                        createBackup = backupWindow.CreateBackup;

                        // If the user wants to remember their choice, update preferences
                        if (backupWindow.RememberChoice)
                        {
                            _preferencesService.SetBackupPreference(createBackup);
                        }
                    }
                    else
                    {
                        // User cancelled, stop processing
                        _loadingService.HideLoading();
                        return;
                    }
                }

                var result = await _excelProcessor.ProcessFilesAsync(
                    MagaraoPath,
                    ImpTopupPath,
                    sheetsToProcess,
                    progress,
                    createBackup);

                // Update UI with results
                DuplicatesFound = result.TotalDuplicates;
                NonDuplicatesFound = result.TotalNonDuplicates;
                SaveLocation = result.SaveLocation;
                InvalidFarmAreasFound = result.InvalidFarmAreaRecords.Count;
                _invalidFarmAreaRecords = result.InvalidFarmAreaRecords;

                // Calculate total endorsed (should be duplicates minus invalid)
                TotalEndorsed = DuplicatesFound - InvalidFarmAreasFound;

                // Set total highlighted RSBSA
                TotalHighlightedRsbsa = DuplicatesFound;

                // Update duplicates list
                _duplicatesList.Clear();
                foreach (var duplicate in result.DuplicateList)
                {
                    _duplicatesList.Add(duplicate);
                }
                _currentPage = 1; // Reset to first page
                UpdateCurrentPageItems(); // Make sure this line is called to update the displayed items

                // Store non-duplicates
                _nonDuplicatesBySheet = result.NonDuplicatesBySheet;

                // Enable buttons
                CanOpenImpTopupFile = true;
                CanOpenRffaFile = true;
                CanViewNonDuplicates = _nonDuplicatesBySheet.Count > 0;
                CanViewInvalidFarmAreas = _invalidFarmAreaRecords.Count > 0;
                CanExportBatch = true;

                // Update status
                StatusMessage = $"Processing complete! Found {result.TotalDuplicates} duplicates, {result.TotalNonDuplicates} non-duplicates, {_invalidFarmAreaRecords.Count} invalid farm areas, {TotalEndorsed} endorsed.";

                // Show success message
                var sheetsStr = string.Join(", ", sheetsToProcess);
                var message = $"Processing complete!\n\n" +
                             $"Processed sheets: {sheetsStr}\n" +
                             $"Found {result.TotalDuplicates} total duplicates.\n";

                if (_nonDuplicatesBySheet.Count > 0)
                {
                    message += $"Found {result.TotalNonDuplicates} total non-duplicate RSBSA numbers.\n";
                }

                if (_invalidFarmAreaRecords.Count > 0)
                {
                    message += $"Found {_invalidFarmAreaRecords.Count} records with farm area outside valid range.\n" +
                              $"Click 'View Invalid Farm Areas' to see details.\n";
                }

                message += $"Total Endorsed: {TotalEndorsed} (Duplicates minus Invalid Farm Areas)\n";
                message += $"Total Highlighted RSBSA in RFFA: {TotalHighlightedRsbsa}\n";

                message += "\nThe files have been updated:";
                message += "\n- RFFA file: Added Duplicates sheet and highlighted duplicate rows";
                message += "\n- IMP Topup file: Highlighted duplicate entries";

                if (createBackup)
                {
                    message += $"\n\nBackup files were created in: {result.SaveLocation}";
                }

                MessageBox.Show(message, "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                // Show invalid farm areas window if any were found
                if (_invalidFarmAreaRecords.Count > 0)
                {
                    ViewInvalidFarmAreas();
                }

                // Hide loading overlay when done
                _loadingService.HideLoading();
            }
            catch (Exception ex)
            {
                _loadingService.HideLoading();
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusMessage = "Error occurred during processing.";
            }
            finally
            {
                OnPropertyChanged(nameof(CanProcess));
            }
        }

        private void OpenImpTopupFile()
        {
            if (string.IsNullOrEmpty(ImpTopupPath) || !File.Exists(ImpTopupPath))
            {
                MessageBox.Show("IMP Topup file not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo(ImpTopupPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenRffaFile()
        {
            if (string.IsNullOrEmpty(MagaraoPath) || !File.Exists(MagaraoPath))
            {
                MessageBox.Show("RFFA file not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo(MagaraoPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ViewInvalidFarmAreas()
        {
            if (_invalidFarmAreaRecords.Count == 0)
            {
                MessageBox.Show("No invalid farm area records found.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                var window = new Views.InvalidFarmAreaWindow(_invalidFarmAreaRecords, MagaraoPath);
                window.Owner = Application.Current.MainWindow;
                window.ShowDialog();

                // After viewing the invalid farm areas, update the status
                int unhighlightedCount = _invalidFarmAreaRecords.Count(r => !r.IsHighlighted);
                if (unhighlightedCount > 0)
                {
                    StatusMessage = $"Processed {unhighlightedCount} of {_invalidFarmAreaRecords.Count} invalid farm area records.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error displaying invalid farm areas: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ViewNonDuplicates()
        {
            try
            {
                if (_nonDuplicatesBySheet == null || _nonDuplicatesBySheet.Count == 0)
                {
                    MessageBox.Show("No non-duplicate RSBSA numbers to display.",
                                   "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Create the window with a try-catch block for safety
                try
                {
                    var window = new Views.NonDuplicatesWindow();
                    window.Owner = Application.Current.MainWindow;

                    // Initialize the window with data before showing it
                    window.Initialize(_nonDuplicatesBySheet);

                    // Show the window
                    window.ShowDialog();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error creating non-duplicates window: {ex.Message}",
                                   "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error displaying non-duplicates: {ex.Message}",
                               "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void ExportHistory()
        {
            if (History.Count == 0)
            {
                MessageBox.Show("No processing history to export.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "Export History",
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                DefaultExt = ".xlsx",
                FileName = "RFFA_Processing_History.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    await _historyService.ExportHistoryToExcel(dialog.FileName);
                    MessageBox.Show($"Processing history exported to:\n{dialog.FileName}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to export history: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void DeleteSelectedHistory()
        {
            if (_selectedHistoryItems.Count == 0)
            {
                MessageBox.Show("No items selected. Please select items to delete.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var result = MessageBox.Show(
                $"Are you sure you want to delete {_selectedHistoryItems.Count} selected item(s)?",
                "Confirm Deletion",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                foreach (var record in _selectedHistoryItems.ToList())
                {
                    _historyService.DeleteRecord(record);
                }

                MessageBox.Show($"{_selectedHistoryItems.Count} history entries deleted.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                _selectedHistoryItems.Clear();

                // Re-evaluate CanExecute
                CommandManager.InvalidateRequerySuggested();
            }
        }

        private void ClearAllHistory()
        {
            if (History.Count == 0)
            {
                MessageBox.Show("History is already empty.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var result = MessageBox.Show(
                $"Are you sure you want to clear all {History.Count} history entries?",
                "Confirm Deletion",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _historyService.ClearHistory();
                MessageBox.Show("All history entries cleared.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}