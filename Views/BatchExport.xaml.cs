using RffaDataComparisonTool.ViewModels;
using System.Windows;

namespace RffaDataComparisonTool.Views
{
    public partial class BatchExportWindow : Window
    {
        private readonly BatchExportViewModel _viewModel;

        public BatchExportWindow(string rffaPath, string impTopupPath, string saveLocation)
        {
            InitializeComponent();

            _viewModel = new BatchExportViewModel(rffaPath, impTopupPath, saveLocation);
            DataContext = _viewModel;

            // Handle ViewModel events
            _viewModel.CloseRequested += (sender, args) =>
            {
                DialogResult = args.Success;
                Close();
            };
        }

        // Property to get the generated batch file path after successful export
        public string GeneratedBatchFilePath => _viewModel.GeneratedBatchFilePath;
    }
}