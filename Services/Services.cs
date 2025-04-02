using RffaDataComparisonTool.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Collections.ObjectModel;

namespace RffaDataComparisonTool.Services
{
    /// <summary>
    /// Service for processing Excel files
    /// </summary>
    public class ExcelProcessorService
    {
        // Event to notify about processing progress
        public event Action<string> LogUpdated;
        private readonly HistoryService _historyService;
        private readonly UserPreferencesService _preferencesService;

        public ExcelProcessorService(HistoryService historyService, UserPreferencesService preferencesService)
        {
            _historyService = historyService;
            _preferencesService = preferencesService;
            // Enable EPPlus non-commercial use
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Get all visible sheets from the MAGARAO-RFFA file
        /// </summary>
        public List<SheetInfo> GetVisibleSheets(string filePath)
        {
            var sheets = new List<SheetInfo>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    // Different versions of EPPlus use different properties to check visibility
                    // Try the most common ones
                    bool isVisible = true;

                    // Try to check if the worksheet is hidden using various properties
                    // that exist in different EPPlus versions
                    try
                    {
                        // For EPPlus 4.x
                        var hiddenProperty = worksheet.GetType().GetProperty("Hidden");
                        if (hiddenProperty != null)
                        {
                            var hiddenValue = (int)hiddenProperty.GetValue(worksheet);
                            isVisible = hiddenValue == 0; // 0 means visible
                        }
                        // For EPPlus 5.x and newer
                        else
                        {
                            var visibleProperty = worksheet.GetType().GetProperty("Visibility");
                            if (visibleProperty != null)
                            {
                                var visibilityValue = visibleProperty.GetValue(worksheet);
                                isVisible = visibilityValue.ToString() == "Visible";
                            }
                        }
                    }
                    catch
                    {
                        // If any error occurs, assume the sheet is visible
                        isVisible = true;
                    }

                    if (isVisible)
                    {
                        sheets.Add(new SheetInfo(worksheet.Name));
                    }
                }
            }

            return sheets;
        }

        /// <summary>
        /// Helper method to find a column index by header name
        /// </summary>
        public static int FindColumnIndex(ExcelWorksheet worksheet, string headerName)
        {
            int columns = worksheet.Dimension.Columns;
            for (int col = 1; col <= columns; col++)
            {
                var cellValue = worksheet.Cells[1, col].Value;
                if (cellValue != null && cellValue.ToString().Trim().Equals(headerName, StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }
            return -1; // Column not found
        }

        /// <summary>
        /// Helper method to open an Excel file
        /// </summary>
        public static ExcelPackage OpenExcelFile(string filePath)
        {
            return new ExcelPackage(new FileInfo(filePath));
        }

        /// <summary>
        /// Process files to identify duplicates and non-duplicates
        /// </summary>
        public async Task<ProcessingResult> ProcessFilesAsync(
            string magaraoPath,
            string impTopupPath,
            List<string> selectedSheets,
            IProgress<string> progress,
            bool createBackup = true)
        {
            var result = new ProcessingResult();

            try
            {
                // Log start of processing
                progress.Report("Starting processing...");

                // CRITICAL: Create backups FIRST, before any file operations
                if (createBackup)
                {
                    try
                    {
                        // Create backup for RFFA file FIRST
                        string rffaDir = Path.GetDirectoryName(magaraoPath);
                        string rffaFileName = Path.GetFileName(magaraoPath);
                        string rffaBackupPath = Path.Combine(rffaDir, "BACKUP_" + rffaFileName);

                        // Use File.Copy to create the backup
                        File.Copy(magaraoPath, rffaBackupPath, true);
                        progress.Report($"Created backup of RFFA file: BACKUP_{rffaFileName}");

                        // Create backup for IMP Topup file
                        string impDir = Path.GetDirectoryName(impTopupPath);
                        string impFileName = Path.GetFileName(impTopupPath);
                        string impBackupPath = Path.Combine(impDir, "BACKUP_" + impFileName);

                        // Use File.Copy to create the backup
                        File.Copy(impTopupPath, impBackupPath, true);
                        progress.Report($"Created backup of IMP Topup file: BACKUP_{impFileName}");

                        // Store backup file names in the result
                        result.BackupFileName = "BACKUP_" + rffaFileName;
                    }
                    catch (Exception ex)
                    {
                        progress.Report($"Warning: Failed to create backup: {ex.Message}");
                        // Continue processing even if backup fails, but warn the user
                    }
                }
                else
                {
                    progress.Report("Proceeding without creating backups (as per user preference).");
                }

                // Step 1: Read all values from IMP Topup file
                progress.Report("Reading IMP Topup file...");
                var impTopupValues = new HashSet<string>();

                using (var package = new ExcelPackage(new FileInfo(impTopupPath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // First worksheet

                    int rows = worksheet.Dimension.Rows;
                    for (int row = 2; row <= rows; row++) // Skip header row
                    {
                        var cellValue = worksheet.Cells[row, 1].Value; // Column A
                        if (cellValue != null)
                        {
                            string cellValueStr = cellValue.ToString().Trim();
                            impTopupValues.Add(cellValueStr);
                        }
                    }
                }

                progress.Report($"Found {impTopupValues.Count} entries in IMP Topup file");

                // Process each selected sheet
                var allDuplicates = new List<string>();
                var nonDuplicatesBySheet = new Dictionary<string, List<string>>();
                var invalidFarmAreaRecords = new List<FarmAreaRecord>();

                // Now process the RFFA file AFTER backup is created
                using (var magaraoPackage = new ExcelPackage(new FileInfo(magaraoPath)))
                {
                    // Create a new sheet for duplicates with a unique name
                    string duplicatesSheetName = "Duplicates";
                    int counter = 1;

                    // Check if sheet exists and create a unique name by appending a number
                    while (magaraoPackage.Workbook.Worksheets.Any(ws => ws.Name.Equals(duplicatesSheetName, StringComparison.OrdinalIgnoreCase)))
                    {
                        duplicatesSheetName = $"Duplicates_{counter++}";
                    }

                    var duplicatesSheet = magaraoPackage.Workbook.Worksheets.Add(duplicatesSheetName);
                    int duplicatesRowIndex = 1;

                    Dictionary<string, int> columnHeaderMap = new Dictionary<string, int>();

                    foreach (var sheetName in selectedSheets)
                    {
                        progress.Report($"\nProcessing sheet: {sheetName}");

                        var worksheet = magaraoPackage.Workbook.Worksheets[sheetName];
                        if (worksheet == null)
                        {
                            progress.Report($"Warning: Sheet '{sheetName}' not found. Skipping.");
                            continue;
                        }

                        // Find column indices
                        int rsbsaColumn = FindColumnIndex(worksheet, "RSBSA NUMBER");
                        int farmAreaColumn = FindColumnIndex(worksheet, "TOTAL FARM AREA (Ha)");
                        int lastNameColumn = FindColumnIndex(worksheet, "LAST NAME");
                        int firstNameColumn = FindColumnIndex(worksheet, "FIRST NAME");
                        int middleNameColumn = FindColumnIndex(worksheet, "MIDDLE NAME");

                        if (rsbsaColumn == -1)
                        {
                            progress.Report($"Warning: RSBSA NUMBER column not found in sheet '{sheetName}'. Searching for 'RSBSA' instead.");
                            rsbsaColumn = FindColumnIndex(worksheet, "RSBSA");
                            if (rsbsaColumn == -1)
                            {
                                progress.Report($"Warning: RSBSA column also not found in sheet '{sheetName}'. Using column B as fallback.");
                                rsbsaColumn = 2; // Fallback to column B
                            }
                        }

                        if (farmAreaColumn == -1)
                        {
                            progress.Report($"Warning: TOTAL FARM AREA (Ha) column not found in sheet '{sheetName}'. Checking alternative names...");
                            farmAreaColumn = FindColumnIndex(worksheet, "FARM AREA");
                            if (farmAreaColumn == -1)
                            {
                                farmAreaColumn = FindColumnIndex(worksheet, "FARM SIZE");
                                if (farmAreaColumn == -1)
                                {
                                    progress.Report($"Warning: Could not find farm area column in sheet '{sheetName}'. Farm area validation will be skipped.");
                                }
                            }
                        }

                        // Store column headers for duplicates sheet
                        if (duplicatesRowIndex == 1)
                        {
                            int columns = worksheet.Dimension.Columns;
                            for (int col = 1; col <= columns; col++)
                            {
                                var headerCell = worksheet.Cells[1, col].Value;
                                if (headerCell != null)
                                {
                                    duplicatesSheet.Cells[1, col].Value = headerCell.ToString();

                                    // Store header index for quick reference
                                    if (!columnHeaderMap.ContainsKey(headerCell.ToString()))
                                    {
                                        columnHeaderMap.Add(headerCell.ToString(), col);
                                    }
                                }
                            }
                            duplicatesRowIndex++;
                        }

                        // Read RSBSA references and process rows
                        var rsbsaRefs = new HashSet<string>();
                        var sheetDuplicates = new List<string>();
                        var sheetNonDuplicates = new List<string>();
                        int rows = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rows; row++) // Skip header row
                        {
                            var cellValue = worksheet.Cells[row, rsbsaColumn].Value; // RSBSA column
                            if (cellValue != null)
                            {
                                string rsbsaNumber = cellValue.ToString().Trim();
                                if (!string.IsNullOrEmpty(rsbsaNumber))
                                {
                                    bool isDuplicate = impTopupValues.Contains(rsbsaNumber);

                                    // Add to appropriate list
                                    if (isDuplicate)
                                    {
                                        sheetDuplicates.Add(rsbsaNumber);

                                        // Highlight the row in the RFFA file
                                        int lastColumn = worksheet.Dimension.Columns;
                                        worksheet.Cells[row, 1, row, lastColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        worksheet.Cells[row, 1, row, lastColumn].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 249, 172)); // Light yellow

                                        // Copy row to duplicates sheet
                                        for (int col = 1; col <= lastColumn; col++)
                                        {
                                            duplicatesSheet.Cells[duplicatesRowIndex, col].Value = worksheet.Cells[row, col].Value;
                                        }
                                        duplicatesRowIndex++;
                                    }
                                    else
                                    {
                                        sheetNonDuplicates.Add(rsbsaNumber);
                                    }

                                    // Check farm area if the column was found
                                    if (farmAreaColumn != -1)
                                    {
                                        var farmAreaCell = worksheet.Cells[row, farmAreaColumn].Value;
                                        if (farmAreaCell != null)
                                        {
                                            double farmArea;
                                            if (double.TryParse(farmAreaCell.ToString().Replace(',', '.'),
                                                                System.Globalization.NumberStyles.Any,
                                                                System.Globalization.CultureInfo.InvariantCulture,
                                                                out farmArea))
                                            {
                                                if (farmArea == 0 || farmArea < 0 || farmArea > 2.0)
                                                {
                                                    string lastName = lastNameColumn != -1 ?
                                                        (worksheet.Cells[row, lastNameColumn].Value?.ToString() ?? "") : "";
                                                    string firstName = firstNameColumn != -1 ?
                                                        (worksheet.Cells[row, firstNameColumn].Value?.ToString() ?? "") : "";
                                                    string middleName = middleNameColumn != -1 ?
                                                        (worksheet.Cells[row, middleNameColumn].Value?.ToString() ?? "") : "";

                                                    invalidFarmAreaRecords.Add(new FarmAreaRecord
                                                    {
                                                        SheetName = sheetName,
                                                        RowNumber = row,
                                                        RsbsaNumber = rsbsaNumber,
                                                        LastName = lastName,
                                                        FirstName = firstName,
                                                        MiddleName = middleName,
                                                        FarmArea = farmArea,
                                                        IsHighlighted = true
                                                    });

                                                    // Highlight the farm area cell
                                                    worksheet.Cells[row, farmAreaColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                    worksheet.Cells[row, farmAreaColumn].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 153, 153)); // Light red
                                                }
                                            }
                                        }
                                    }

                                    rsbsaRefs.Add(rsbsaNumber);
                                }
                            }
                        }

                        // Add non-duplicates for this sheet
                        if (sheetNonDuplicates.Count > 0)
                        {
                            nonDuplicatesBySheet[sheetName] = sheetNonDuplicates;
                        }

                        // Update totals
                        allDuplicates.AddRange(sheetDuplicates);

                        // Add to history
                        _historyService.AddRecord(new ProcessingRecord(
                            DateTime.Now,
                            sheetName,
                            sheetDuplicates.Count,
                            sheetNonDuplicates.Count
                        ));

                        progress.Report($"Sheet '{sheetName}': Found {sheetDuplicates.Count} duplicates, {sheetNonDuplicates.Count} non-duplicates");
                        if (farmAreaColumn != -1)
                        {
                            int invalidCount = invalidFarmAreaRecords.Count(r => r.SheetName == sheetName);
                            progress.Report($"Found {invalidCount} records with farm area outside range 0-2.0 Ha");
                        }
                    }

                    // Format the duplicates sheet headers
                    using (var headerRange = duplicatesSheet.Cells[1, 1, 1, duplicatesSheet.Dimension?.Columns ?? 1])
                    {
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }

                    // Save the changes to RFFA file
                    magaraoPackage.Save();
                    progress.Report("Saved changes to RFFA file with highlighted duplicates and new Duplicates sheet");
                }

                // Highlight duplicates in the IMP Topup file
                progress.Report("\nHighlighting duplicates in IMP Topup file...");

                int highlightedCount = 0;

                using (var package = new ExcelPackage(new FileInfo(impTopupPath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // First worksheet

                    // Define the turquoise fill
                    var fill = worksheet.Cells["A1"].Style.Fill;
                    fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(119, 223, 216)); // Turquoise

                    int rows = worksheet.Dimension.Rows;
                    for (int row = 2; row <= rows; row++) // Skip header row
                    {
                        var cellValue = worksheet.Cells[row, 1].Value; // Column A
                        if (cellValue != null)
                        {
                            string cellValueStr = cellValue.ToString().Trim();

                            if (allDuplicates.Contains(cellValueStr))
                            {
                                highlightedCount++;

                                // Highlight the entire row
                                int lastColumn = Math.Min(20, worksheet.Dimension.Columns);
                                worksheet.Cells[row, 1, row, lastColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                worksheet.Cells[row, 1, row, lastColumn].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(119, 223, 216));
                            }
                        }
                    }

                    // Save the changes
                    package.Save();
                }

                progress.Report($"Highlighted {highlightedCount} duplicate entries in IMP Topup file");

                // Store actual highlighted count in result for UI display
                result.HighlightedRowsInImpTopup = highlightedCount;

                // Prepare the result
                result.TotalDuplicates = allDuplicates.Count;
                result.TotalNonDuplicates = nonDuplicatesBySheet.Values.Sum(list => list.Count);
                result.DuplicateList = allDuplicates.Distinct().ToList();
                result.NonDuplicatesBySheet = nonDuplicatesBySheet;
                result.SaveLocation = Path.GetDirectoryName(magaraoPath);
                result.InvalidFarmAreaRecords = invalidFarmAreaRecords;

                return result;
            }
            catch (Exception ex)
            {
                progress.Report($"Error: {ex.Message}");
                progress.Report(ex.StackTrace);
                throw;
            }
        }

        /// <summary>
        /// Remove highlight from a specific row with invalid farm area
        /// </summary>
        public async Task<bool> RemoveHighlightAsync(string filePath, FarmAreaRecord record)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[record.SheetName];
                    if (worksheet != null)
                    {
                        // Find farm area column
                        int farmAreaColumn = FindColumnIndex(worksheet, "TOTAL FARM AREA (Ha)");
                        if (farmAreaColumn == -1)
                        {
                            farmAreaColumn = FindColumnIndex(worksheet, "FARM AREA");
                            if (farmAreaColumn == -1)
                            {
                                farmAreaColumn = FindColumnIndex(worksheet, "FARM SIZE");
                            }
                        }

                        if (farmAreaColumn != -1)
                        {
                            // Remove highlighting from the farm area cell
                            worksheet.Cells[record.RowNumber, farmAreaColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.None;
                        }

                        // Save changes
                        package.Save();
                        return true;
                    }
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Export invalid farm area records to Excel
        /// </summary>
        public async Task<string> ExportInvalidFarmAreasToExcel(List<FarmAreaRecord> records, string savePath)
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Invalid Farm Areas");

                    // Add headers
                    worksheet.Cells[1, 1].Value = "Sheet Name";
                    worksheet.Cells[1, 2].Value = "RSBSA Number";
                    worksheet.Cells[1, 3].Value = "Last Name";
                    worksheet.Cells[1, 4].Value = "First Name";
                    worksheet.Cells[1, 5].Value = "Middle Name";
                    worksheet.Cells[1, 6].Value = "Farm Area (Ha)";
                    worksheet.Cells[1, 7].Value = "Row Number";

                    // Style the headers
                    using (var range = worksheet.Cells[1, 1, 1, 7])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }

                    // Add data
                    for (int i = 0; i < records.Count; i++)
                    {
                        var record = records[i];
                        int row = i + 2;

                        worksheet.Cells[row, 1].Value = record.SheetName;
                        worksheet.Cells[row, 2].Value = record.RsbsaNumber;
                        worksheet.Cells[row, 3].Value = record.LastName;
                        worksheet.Cells[row, 4].Value = record.FirstName;
                        worksheet.Cells[row, 5].Value = record.MiddleName;
                        worksheet.Cells[row, 6].Value = record.FarmArea;
                        worksheet.Cells[row, 7].Value = record.RowNumber;

                        // Highlight rows with farm areas outside the valid range
                        if (record.FarmArea > 2.0 || record.FarmArea < 0.1)
                        {
                            var color = record.FarmArea > 2.0
                                ? System.Drawing.Color.FromArgb(255, 235, 156) // Light orange for values > 2.0
                                : System.Drawing.Color.FromArgb(255, 199, 206); // Light red for values < 0.1

                            worksheet.Cells[row, 1, row, 7].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[row, 1, row, 7].Style.Fill.BackgroundColor.SetColor(color);
                        }
                    }

                    // Format cells
                    worksheet.Cells[2, 6, records.Count + 1, 6].Style.Numberformat.Format = "0.00";
                    worksheet.Cells[1, 1, records.Count + 1, 7].AutoFitColumns();

                    // Save the file
                    string fileName = "Invalid_Farm_Areas.xlsx";
                    string filePath = Path.Combine(savePath, fileName);
                    await package.SaveAsAsync(new FileInfo(filePath));
                    return filePath;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to export invalid farm areas: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Service for managing processing history
    /// </summary>
    public class HistoryService
    {
        private readonly ObservableCollection<ProcessingRecord> _history = new ObservableCollection<ProcessingRecord>();

        public ObservableCollection<ProcessingRecord> History => _history;

        // Clear history before adding new records for a new processing operation
        public void ClearHistoryForNewProcess()
        {
            _history.Clear();
        }

        public void AddRecord(ProcessingRecord record)
        {
            _history.Add(record);
        }

        public void DeleteRecord(ProcessingRecord record)
        {
            if (_history.Contains(record))
            {
                _history.Remove(record);
            }
        }

        public void ClearHistory()
        {
            _history.Clear();
        }

        public async Task ExportHistoryToExcel(string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("History");

                // Add headers
                worksheet.Cells[1, 1].Value = "Date/Time";
                worksheet.Cells[1, 2].Value = "Sheet";
                worksheet.Cells[1, 3].Value = "Duplicates";
                worksheet.Cells[1, 4].Value = "Non-Duplicates";

                // Add data
                for (int i = 0; i < _history.Count; i++)
                {
                    var record = _history[i];
                    int row = i + 2;

                    worksheet.Cells[row, 1].Value = record.DateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    worksheet.Cells[row, 2].Value = record.SheetName;
                    worksheet.Cells[row, 3].Value = record.DuplicateCount;
                    worksheet.Cells[row, 4].Value = record.NonDuplicateCount;
                }

                // Style the headers
                using (var range = worksheet.Cells[1, 1, 1, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // Save the file
                await package.SaveAsAsync(new FileInfo(filePath));
            }
        }
    }

    /// <summary>
    /// Interface for loading service
    /// </summary>
    public interface ILoadingService
    {
        void ShowLoading(string message);
        void HideLoading();
    }

    /// <summary>
    /// Simple loading service implementation
    /// </summary>
    public class LoadingService : ILoadingService
    {
        private readonly Action<string> _showLoadingAction;
        private readonly Action _hideLoadingAction;

        public LoadingService(Action<string> showLoadingAction, Action hideLoadingAction)
        {
            _showLoadingAction = showLoadingAction;
            _hideLoadingAction = hideLoadingAction;
        }

        public void ShowLoading(string message)
        {
            _showLoadingAction?.Invoke(message);
        }

        public void HideLoading()
        {
            _hideLoadingAction?.Invoke();
        }
    }

    /// <summary>
    /// Service for managing user preferences
    /// </summary>
    public class UserPreferencesService
    {
        private const string PREFERENCE_FILE = "preferences.json";
        private UserPreferences _preferences;

        public UserPreferencesService()
        {
            LoadPreferences();
        }

        public bool AlwaysCreateBackup => _preferences.AlwaysCreateBackup;

        public void SetBackupPreference(bool alwaysCreateBackup)
        {
            _preferences.AlwaysCreateBackup = alwaysCreateBackup;
            SavePreferences();
        }

        private void LoadPreferences()
        {
            try
            {
                if (File.Exists(PREFERENCE_FILE))
                {
                    string json = File.ReadAllText(PREFERENCE_FILE);
                    _preferences = System.Text.Json.JsonSerializer.Deserialize<UserPreferences>(json);
                }
                else
                {
                    _preferences = new UserPreferences { AlwaysCreateBackup = true };
                    SavePreferences();
                }
            }
            catch
            {
                _preferences = new UserPreferences { AlwaysCreateBackup = true };
            }
        }

        private void SavePreferences()
        {
            try
            {
                string json = System.Text.Json.JsonSerializer.Serialize(_preferences);
                File.WriteAllText(PREFERENCE_FILE, json);
            }
            catch
            {
                // If saving fails, silently continue
            }
        }
    }
}