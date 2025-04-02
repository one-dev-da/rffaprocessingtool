using RffaDataComparisonTool.ViewModels;
using System;
using System.Collections.Generic;
using System.Windows;
using RffaDataComparisonTool.Models;

namespace RffaDataComparisonTool.Views
{
    public partial class InvalidFarmAreaWindow : Window
    {
        private readonly InvalidFarmAreaViewModel _viewModel;

        public InvalidFarmAreaWindow(List<FarmAreaRecord> records, string rffaPath)
        {
            InitializeComponent();

            // Initialize ViewModel
            _viewModel = new InvalidFarmAreaViewModel(records, rffaPath);
            DataContext = _viewModel;

            // Configure the ViewModel's DialogResult to close the window
            _viewModel.PropertyChanged += (sender, e) =>
            {
                if (e.PropertyName == nameof(_viewModel.DialogResult) && _viewModel.DialogResult.HasValue)
                {
                    DialogResult = _viewModel.DialogResult;
                }
            };

            // Ensure we don't minimize main window when closing
            this.Closing += (s, e) =>
            {
                this.Owner = null;
            };
        }
    }
}