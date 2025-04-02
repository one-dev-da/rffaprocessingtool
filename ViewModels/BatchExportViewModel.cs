using Microsoft.Win32;
using RffaDataComparisonTool.Helpers;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using OfficeOpenXml;
using System.Windows.Media;
using System.Collections.Generic;

namespace RffaDataComparisonTool.ViewModels
{
    public class BatchExportViewModel : ObservableObject
    {
        private readonly string _rffaFilePath;
        private readonly string _impTopupPath;
        private readonly string _saveLocation;

        // Properties for binding
        private string _batchNumber;
        public string BatchNumber
        {
            get => _batchNumber;
            set
            {
                if (SetProperty(ref _batchNumber, value, nameof(BatchNumber)))
                {
                    OnPropertyChanged(nameof(CanExport));
                }
            }
        }

        private string _province;
        public string Province
        {
            get => _province;
            set
            {
                if (SetProperty(ref _province, value, nameof(Province)))
                {
                    OnPropertyChanged(nameof(CanExport));
                }
            }
        }

        private string _municipality;
        public string Municipality
        {
            get => _municipality;
            set
            {
                if (SetProperty(ref _municipality, value, nameof(Municipality)))
                {
                    OnPropertyChanged(nameof(CanExport));
                }
            }
        }

        private bool _updateExisting;
        public bool UpdateExisting
        {
            get => _updateExisting;
            set
            {
                if (SetProperty(ref _updateExisting, value, nameof(UpdateExisting)))
                {
                    OnPropertyChanged(nameof(ExistingFileSelectVisibility));
                    OnPropertyChanged(nameof(CanExport));
                }
            }
        }

        private string _existingFilePath;
        public string ExistingFilePath
        {
            get => _existingFilePath;
            set
            {
                if (SetProperty(ref _existingFilePath, value, nameof(ExistingFilePath)))
                {
                    OnPropertyChanged(nameof(CanExport));
                }
            }
        }

        public string GeneratedBatchFilePath { get; private set; }

        // Helper properties
        public Visibility ExistingFileSelectVisibility => UpdateExisting ? Visibility.Visible : Visibility.Collapsed;

        public bool CanExport =>
    // When updating existing file, Municipality and ExistingFilePath are required
    UpdateExisting
        ? !string.IsNullOrWhiteSpace(Municipality) && !string.IsNullOrWhiteSpace(ExistingFilePath)
        // When creating new file, all fields are required
        : !string.IsNullOrWhiteSpace(BatchNumber) &&
          !string.IsNullOrWhiteSpace(Province) &&
          !string.IsNullOrWhiteSpace(Municipality);

        // Commands
        public ICommand ExportCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand BrowseExistingFileCommand { get; }

        // Event for closing the window
        public event EventHandler<CloseRequestedEventArgs> CloseRequested;

        public BatchExportViewModel(string rffaPath, string impTopupPath, string saveLocation)
        {
            _rffaFilePath = rffaPath;
            _impTopupPath = impTopupPath;
            _saveLocation = saveLocation;

            // Initialize commands
            ExportCommand = new RelayCommand(_ => Export(), _ => CanExport);
            CancelCommand = new RelayCommand(_ => Cancel());
            BrowseExistingFileCommand = new RelayCommand(_ => BrowseExistingFile());
        }

        private void BrowseExistingFile()
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Existing Batch File",
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                CheckFileExists = true,
                InitialDirectory = _saveLocation
            };

            if (dialog.ShowDialog() == true)
            {
                ExistingFilePath = dialog.FileName;
                // Ensure binding updates are triggered
                OnPropertyChanged(nameof(ExistingFilePath));
                OnPropertyChanged(nameof(CanExport));
            }
        }

        // Update the Export method in BatchExportViewModel.cs:

        private async void Export()
        {
            try
            {
                // Only validate batch number when creating a new file
                if (!UpdateExisting)
                {
                    // Validate batch number (should be numeric)
                    if (!int.TryParse(BatchNumber, out _))
                    {
                        MessageBox.Show("Batch number must be numeric.", "Validation Error",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }

                // Generate file name and path
                string batchFilePath;

                if (UpdateExisting && !string.IsNullOrEmpty(ExistingFilePath))
                {
                    // Use the existing file path directly
                    batchFilePath = ExistingFilePath;

                    // Make sure the file exists and is accessible
                    if (!File.Exists(batchFilePath))
                    {
                        MessageBox.Show($"The selected file does not exist:\n{batchFilePath}",
                            "File Not Found", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // Check if the file is accessible (not locked)
                    try
                    {
                        using (var fs = new FileStream(batchFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                        {
                            // Just testing if we can open it
                        }
                    }
                    catch (IOException)
                    {
                        MessageBox.Show("The selected file is currently open in another program. " +
                            "Please close it before updating.", "File Locked", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }
                else
                {
                    // Clean province name (remove special characters)
                    string cleanProvince = string.Join("_", Province.Split(Path.GetInvalidFileNameChars()));

                    // Create a new file path
                    string batchFileName = $"BATCH_{BatchNumber}_{cleanProvince}.xlsx";
                    batchFilePath = Path.Combine(_saveLocation, batchFileName);

                    // Check if directory exists, create it if it doesn't
                    string directoryPath = Path.GetDirectoryName(batchFilePath);
                    if (!Directory.Exists(directoryPath))
                    {
                        try
                        {
                            Directory.CreateDirectory(directoryPath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Failed to create directory: {ex.Message}",
                                "Directory Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                    }

                    // Check if file already exists - ask for confirmation to overwrite
                    if (File.Exists(batchFilePath))
                    {
                        var result = MessageBox.Show($"The file already exists:\n{batchFilePath}\n\nDo you want to overwrite it?",
                            "File Exists", MessageBoxButton.YesNo, MessageBoxImage.Question);

                        if (result != MessageBoxResult.Yes)
                            return;

                        // Try to delete existing file to avoid issues with locked files
                        try
                        {
                            File.Delete(batchFilePath);
                        }
                        catch (IOException)
                        {
                            MessageBox.Show("The file is currently open in another program. " +
                                "Please close it before overwriting.", "File Locked", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Failed to overwrite existing file: {ex.Message}",
                                "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                    }
                }

                // Show a simple message that we're exporting
                MessageBox.Show("Exporting batch file. This may take a moment...",
                    "Exporting", MessageBoxButton.OK, MessageBoxImage.Information);

                // Perform the actual export (on a background thread to keep UI responsive)
                bool success = await Task.Run(() => ExportBatchFile(batchFilePath));

                if (success)
                {
                    GeneratedBatchFilePath = batchFilePath;

                    // Use Application.Current.Dispatcher to make sure we're on the UI thread
                    Application.Current.Dispatcher.Invoke(() => {
                        // Close the window with success
                        CloseRequested?.Invoke(this, new CloseRequestedEventArgs { Success = true });
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting batch file: {ex.Message}",
                    "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool ExportBatchFile(string batchFilePath)
        {
            try
            {
                // Enable EPPlus non-commercial use if needed
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Create or open Excel package based on update mode
                using (var package = UpdateExisting && File.Exists(batchFilePath)
                    ? new ExcelPackage(new FileInfo(batchFilePath))
                    : new ExcelPackage())
                {
                    // Clean municipality name for the sheet name (remove invalid characters)
                    string cleanMunicipality = string.Join("_", Municipality.Split(Path.GetInvalidFileNameChars()));

                    // Generate a unique sheet name based on the municipality
                    string sheetName = cleanMunicipality;

                    // Make sure sheet name is unique by appending a number if needed
                    int counter = 1;
                    string baseSheetName = sheetName;
                    while (package.Workbook.Worksheets[sheetName] != null)
                    {
                        sheetName = $"{baseSheetName}_{counter++}";
                    }

                    // Create new sheet
                    var worksheet = package.Workbook.Worksheets.Add(sheetName);

                    // First step: Read all values from RFFA file to identify duplicates
                    var rffaDuplicates = new Dictionary<string, List<object[]>>();

                    // Open the RFFA file to find duplicates
                    using (var rffaPackage = new ExcelPackage(new FileInfo(_rffaFilePath)))
                    {
                        // Look for the Duplicates sheet first
                        var duplicatesSheet = rffaPackage.Workbook.Worksheets.FirstOrDefault(
                            ws => ws.Name.StartsWith("Duplicates", StringComparison.OrdinalIgnoreCase));

                        if (duplicatesSheet != null && duplicatesSheet.Dimension != null)
                        {
                            // Get the RSBSA column index from the duplicates sheet
                            int rsbsaColumn = FindRsbsaColumnIndex(duplicatesSheet);

                            if (rsbsaColumn > 0)
                            {
                                int rows = duplicatesSheet.Dimension.Rows;
                                int cols = duplicatesSheet.Dimension.Columns;

                                // Read all rows from the Duplicates sheet
                                for (int row = 2; row <= rows; row++) // Skip header
                                {
                                    string rsbsaValue = duplicatesSheet.Cells[row, rsbsaColumn].Value?.ToString();

                                    if (!string.IsNullOrEmpty(rsbsaValue))
                                    {
                                        // Store the entire row data
                                        var rowData = new object[cols];
                                        for (int col = 1; col <= cols; col++)
                                        {
                                            rowData[col - 1] = duplicatesSheet.Cells[row, col].Value;
                                        }

                                        if (!rffaDuplicates.ContainsKey(rsbsaValue))
                                        {
                                            rffaDuplicates.Add(rsbsaValue, new List<object[]>());
                                        }

                                        rffaDuplicates[rsbsaValue].Add(rowData);
                                    }
                                }
                            }
                        }

                        // If no Duplicates sheet, check other sheets for highlighted rows
                        if (rffaDuplicates.Count == 0)
                        {
                            foreach (var sheet in rffaPackage.Workbook.Worksheets)
                            {
                                // Skip Duplicates sheet (already processed) or empty sheets
                                if (sheet.Name.StartsWith("Duplicates", StringComparison.OrdinalIgnoreCase) ||
                                    sheet.Dimension == null)
                                    continue;

                                // Find RSBSA column in this sheet
                                int rsbsaColumn = FindRsbsaColumnIndex(sheet);

                                if (rsbsaColumn > 0)
                                {
                                    int rows = sheet.Dimension.Rows;
                                    int cols = sheet.Dimension.Columns;

                                    // Check each row for highlights
                                    for (int row = 2; row <= rows; row++) // Skip header
                                    {
                                        // Check if this row is highlighted
                                        bool isHighlighted = IsRowHighlighted(sheet, row);

                                        if (isHighlighted)
                                        {
                                            string rsbsaValue = sheet.Cells[row, rsbsaColumn].Value?.ToString();

                                            if (!string.IsNullOrEmpty(rsbsaValue))
                                            {
                                                // Store the entire row data
                                                var rowData = new object[cols];
                                                for (int col = 1; col <= cols; col++)
                                                {
                                                    rowData[col - 1] = sheet.Cells[row, col].Value;
                                                }

                                                if (!rffaDuplicates.ContainsKey(rsbsaValue))
                                                {
                                                    rffaDuplicates.Add(rsbsaValue, new List<object[]>());
                                                }

                                                rffaDuplicates[rsbsaValue].Add(rowData);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Second step: Process the IMP Top-up file to find matches and create the batch export
                    using (var impPackage = new ExcelPackage(new FileInfo(_impTopupPath)))
                    {
                        // We only need to process the first worksheet of the IMP Top-up file
                        var sourceSheet = impPackage.Workbook.Worksheets[0];

                        if (sourceSheet.Dimension == null)
                            throw new Exception("The IMP Top-up file appears to be empty.");

                        int destRow = 1;
                        bool headersAdded = false;

                        int rows = sourceSheet.Dimension.Rows;
                        int cols = sourceSheet.Dimension.Columns;

                        // Find the RSBSA column (typically the first column)
                        int rsbsaColumn = 1; // Default to first column

                        // Try to find the correct column by name
                        for (int col = 1; col <= cols; col++)
                        {
                            var header = sourceSheet.Cells[1, col].Value?.ToString()?.Trim();
                            if (!string.IsNullOrEmpty(header) &&
                                (header.Equals("RSBSA NUMBER", StringComparison.OrdinalIgnoreCase) ||
                                 header.Equals("RSBSA", StringComparison.OrdinalIgnoreCase)))
                            {
                                rsbsaColumn = col;
                                break;
                            }
                        }

                        // Add headers from source
                        for (int col = 1; col <= cols; col++)
                        {
                            var headerValue = sourceSheet.Cells[1, col].Value;
                            if (headerValue != null)
                            {
                                worksheet.Cells[1, col].Value = headerValue;
                                worksheet.Cells[1, col].Style.Font.Bold = true;
                            }
                        }

                        // Add additional metadata columns
                        worksheet.Cells[1, cols + 1].Value = "Export Date";
                        worksheet.Cells[1, cols + 1].Style.Font.Bold = true;

                        worksheet.Cells[1, cols + 2].Value = "Batch Number";
                        worksheet.Cells[1, cols + 2].Style.Font.Bold = true;

                        worksheet.Cells[1, cols + 3].Value = "Province";
                        worksheet.Cells[1, cols + 3].Style.Font.Bold = true;

                        worksheet.Cells[1, cols + 4].Value = "Municipality";
                        worksheet.Cells[1, cols + 4].Style.Font.Bold = true;

                        worksheet.Cells[1, cols + 5].Value = "Processing Date";
                        worksheet.Cells[1, cols + 5].Style.Font.Bold = true;

                        // Format header row
                        var headerRange = worksheet.Cells[1, 1, 1, cols + 5];
                        headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                        destRow++; // Move to row 2 for data

                        // Track how many rows we've added
                        int rowsAdded = 0;
                        HashSet<string> processedRsbsa = new HashSet<string>();

                        // Go through each row in the IMP Top-up file
                        for (int row = 2; row <= rows; row++) // Skip header
                        {
                            string rsbsaValue = sourceSheet.Cells[row, rsbsaColumn].Value?.ToString();

                            // Check if this RSBSA is in our list of duplicates from RFFA file
                            if (!string.IsNullOrEmpty(rsbsaValue) &&
                                rffaDuplicates.ContainsKey(rsbsaValue) &&
                                !processedRsbsa.Contains(rsbsaValue))
                            {
                                // This is a duplicate - add it to the export
                                processedRsbsa.Add(rsbsaValue);

                                // Copy entire row from IMP Top-up
                                for (int col = 1; col <= cols; col++)
                                {
                                    worksheet.Cells[destRow, col].Value = sourceSheet.Cells[row, col].Value;
                                }

                                // Add metadata
                                worksheet.Cells[destRow, cols + 1].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                worksheet.Cells[destRow, cols + 2].Value = BatchNumber;
                                worksheet.Cells[destRow, cols + 3].Value = Province;
                                worksheet.Cells[destRow, cols + 4].Value = Municipality;
                                worksheet.Cells[destRow, cols + 5].Value = DateTime.Now.ToString("yyyy-MM-dd");

                                destRow++;
                                rowsAdded++;
                            }
                        }

                        // Auto-fit columns
                        if (worksheet.Dimension != null)
                        {
                            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                        }

                        // Provide feedback on how many rows were added
                        MessageBox.Show($"Added {rowsAdded} duplicate rows to the batch export.",
                            "Export Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    }

                    // Save the batch file with better error handling
                    try
                    {
                        // First create a temporary file to avoid issues with partially written files
                        string tempFilePath = Path.Combine(
                            Path.GetDirectoryName(batchFilePath),
                            $"temp_{Guid.NewGuid().ToString()}.xlsx");

                        package.SaveAs(new FileInfo(tempFilePath));

                        // If successful, replace or create the final file
                        if (File.Exists(batchFilePath))
                        {
                            // Try to delete existing file first to avoid file in use issues
                            File.Delete(batchFilePath);
                        }

                        // Move the temp file to the final location
                        File.Move(tempFilePath, batchFilePath);

                        return true;
                    }
                    catch (IOException ex)
                    {
                        // Handle file access errors specifically
                        Application.Current.Dispatcher.Invoke(() => {
                            MessageBox.Show($"Cannot save the file because it is being used by another process. " +
                                $"Please close any programs that might be using this file and try again.\n\n" +
                                $"File: {batchFilePath}",
                                "File Access Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        });
                        return false;
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        // Handle permission errors
                        Application.Current.Dispatcher.Invoke(() => {
                            MessageBox.Show($"You don't have permission to save the file in this location. " +
                                $"Try running the application as administrator or choosing a different location.\n\n" +
                                $"File: {batchFilePath}",
                                "Permission Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        });
                        return false;
                    }
                    catch (Exception ex)
                    {
                        // Handle general errors
                        Application.Current.Dispatcher.Invoke(() => {
                            MessageBox.Show($"Error saving batch file: {ex.Message}\n\n" +
                                $"File: {batchFilePath}",
                                "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        });
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating batch file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }            
        }

        // Helper method to find the RSBSA column in a worksheet
        private int FindRsbsaColumnIndex(OfficeOpenXml.ExcelWorksheet sheet)
        {
            if (sheet.Dimension == null)
                return -1;

            int cols = sheet.Dimension.Columns;

            // Look for common RSBSA column names
            for (int col = 1; col <= cols; col++)
            {
                var header = sheet.Cells[1, col].Value?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(header) &&
                    (header.Equals("RSBSA NUMBER", StringComparison.OrdinalIgnoreCase) ||
                     header.Equals("RSBSA", StringComparison.OrdinalIgnoreCase)))
                {
                    return col;
                }
            }

            // If no exact match found, try column B (common location for RSBSA)
            return 2;
        }

        // Helper method to check if a row is highlighted
        private bool IsRowHighlighted(OfficeOpenXml.ExcelWorksheet sheet, int row)
        {
            if (sheet.Dimension == null)
                return false;

            int cols = sheet.Dimension.Columns;

            // Check the first few cells for highlighting
            for (int col = 1; col <= Math.Min(5, cols); col++)
            {
                var cell = sheet.Cells[row, col];
                if (cell.Style.Fill.PatternType == OfficeOpenXml.Style.ExcelFillStyle.Solid)
                {
                    return true;
                }
            }

            return false;
        }

        private void Cancel()
        {
            CloseRequested?.Invoke(this, new CloseRequestedEventArgs { Success = false });
        }
    }

    public class CloseRequestedEventArgs : EventArgs
    {
        public bool Success { get; set; }
    }
}