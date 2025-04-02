using RffaDataComparisonTool.ViewModels;
using System.Windows;

namespace RffaDataComparisonTool.Views
{
    public partial class BackupConfirmWindow : Window
    {
        private readonly BackupConfirmViewModel _viewModel;

        public bool CreateBackup { get; private set; }
        public bool RememberChoice { get; private set; }

        public BackupConfirmWindow()
        {
            InitializeComponent();

            _viewModel = new BackupConfirmViewModel();
            DataContext = _viewModel;

            // Handle ViewModel events
            _viewModel.BackupConfirmed += (sender, args) =>
            {
                CreateBackup = args.CreateBackup;
                RememberChoice = args.RememberChoice;
                DialogResult = true;
            };
        }
    }
}