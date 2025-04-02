using RffaDataComparisonTool.Helpers;
using RffaDataComparisonTool.Models;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;

namespace RffaDataComparisonTool.ViewModels
{
    /// <summary>
    /// ViewModel for sheet selection window
    /// </summary>
    public class SheetSelectionViewModel : ObservableObject
    {
        private ObservableCollection<SheetInfo> _sheets = new ObservableCollection<SheetInfo>();
        public ObservableCollection<SheetInfo> Sheets => _sheets;

        private bool _selectAll;
        public bool SelectAll
        {
            get => _selectAll;
            set
            {
                if (SetProperty(ref _selectAll, value, nameof(SelectAll)))
                {
                    // Update all sheet selections without triggering individual change events
                    foreach (var sheet in _sheets)
                    {
                        sheet.PropertyChanged -= Sheet_PropertyChanged; // Temporarily remove handler
                        sheet.IsSelected = value;
                        sheet.PropertyChanged += Sheet_PropertyChanged; // Restore handler
                    }
                    OnPropertyChanged(nameof(Sheets));
                }
            }
        }

        public List<string> SelectedSheets => _sheets.Where(s => s.IsSelected).Select(s => s.Name).ToList();

        public ICommand OkCommand { get; }
        public ICommand CancelCommand { get; }

        public SheetSelectionViewModel(List<SheetInfo> sheets)
        {
            // Subscribe to property changes before adding to collection
            foreach (var sheet in sheets)
            {
                sheet.PropertyChanged += Sheet_PropertyChanged;
                _sheets.Add(sheet);
            }

            OkCommand = new RelayCommand(_ => DialogResult = true, _ => SelectedSheets.Count > 0);
            CancelCommand = new RelayCommand(_ => DialogResult = false);

            // Initialize SelectAll state based on all sheets being selected
            UpdateSelectAllState();
        }

        private void Sheet_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SheetInfo.IsSelected))
            {
                // Update SelectAll checkbox state when individual selections change
                UpdateSelectAllState();

                // Update OkCommand's CanExecute status
                CommandManager.InvalidateRequerySuggested();
            }
        }

        private void UpdateSelectAllState()
        {
            // Update SelectAll checkbox without triggering its property change event
            bool allSelected = _sheets.Count > 0 && _sheets.All(s => s.IsSelected);
            if (_selectAll != allSelected)
            {
                _selectAll = allSelected;
                OnPropertyChanged(nameof(SelectAll));
            }
        }

        private bool? _dialogResult;
        public bool? DialogResult
        {
            get => _dialogResult;
            set => SetProperty(ref _dialogResult, value, nameof(DialogResult));
        }
    }
}