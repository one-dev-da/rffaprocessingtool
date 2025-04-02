using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace RffaDataComparisonTool.Models
{
    /// <summary>
    /// Model representing a processing history record
    /// </summary>
    public class ProcessingRecord
    {
        public DateTime DateTime { get; set; }
        public string SheetName { get; set; }
        public int DuplicateCount { get; set; }
        public int NonDuplicateCount { get; set; }

        public ProcessingRecord(DateTime dateTime, string sheetName, int duplicateCount, int nonDuplicateCount)
        {
            DateTime = dateTime;
            SheetName = sheetName;
            DuplicateCount = duplicateCount;
            NonDuplicateCount = nonDuplicateCount;
        }
    }

    /// <summary>
    /// Model representing a sheet available for selection
    /// </summary>
    public class SheetInfo : INotifyPropertyChanged
    {
        private string _name;
        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }
        }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        private bool _isVisible;
        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                if (_isVisible != value)
                {
                    _isVisible = value;
                    OnPropertyChanged(nameof(IsVisible));
                }
            }
        }

        public SheetInfo(string name, bool isVisible = true)
        {
            _name = name;
            _isVisible = isVisible;
            _isSelected = false;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Model representing a record with invalid farm area
    /// </summary>
    public class FarmAreaRecord : INotifyPropertyChanged
    {
        public string SheetName { get; set; }
        public int RowNumber { get; set; }
        public string RsbsaNumber { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public double FarmArea { get; set; }

        private bool _isHighlighted;
        public bool IsHighlighted
        {
            get => _isHighlighted;
            set
            {
                if (_isHighlighted != value)
                {
                    _isHighlighted = value;
                    OnPropertyChanged(nameof(IsHighlighted));
                }
            }
        }

        public string FullName => $"{LastName}, {FirstName} {MiddleName}".Trim();

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Model representing user preferences
    /// </summary>
    public class UserPreferences
    {
        public bool AlwaysCreateBackup { get; set; } = true;
    }

    /// <summary>
    /// Model representing the results of processing
    /// </summary>
    public class ProcessingResult
    {
        public int TotalDuplicates { get; set; }
        public int TotalNonDuplicates { get; set; }
        public List<string> DuplicateList { get; set; }
        public Dictionary<string, List<string>> NonDuplicatesBySheet { get; set; }
        public string SaveLocation { get; set; }
        public string BackupFileName { get; set; }
        public List<FarmAreaRecord> InvalidFarmAreaRecords { get; set; }

        // Added property to track the actual number of rows highlighted in IMP Topup file
        public int HighlightedRowsInImpTopup { get; set; }

        public ProcessingResult()
        {
            DuplicateList = new List<string>();
            NonDuplicatesBySheet = new Dictionary<string, List<string>>();
            InvalidFarmAreaRecords = new List<FarmAreaRecord>();
        }
    }
}