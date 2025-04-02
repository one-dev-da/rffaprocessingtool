using RffaDataComparisonTool.Helpers;
using System;
using System.Windows.Input;

namespace RffaDataComparisonTool.ViewModels
{
    public class BackupConfirmViewModel : ObservableObject
    {
        private bool _rememberChoice;
        public bool RememberChoice
        {
            get => _rememberChoice;
            set => SetProperty(ref _rememberChoice, value, nameof(RememberChoice));
        }

        public ICommand CreateBackupCommand { get; }
        public ICommand NoBackupCommand { get; }

        // Event for notifying the window of the user's choice
        public event EventHandler<BackupConfirmationEventArgs> BackupConfirmed;

        public BackupConfirmViewModel()
        {
            CreateBackupCommand = new RelayCommand(_ => ConfirmBackup(true));
            NoBackupCommand = new RelayCommand(_ => ConfirmBackup(false));
        }

        private void ConfirmBackup(bool createBackup)
        {
            // Raise the event with the user's choice
            BackupConfirmed?.Invoke(this, new BackupConfirmationEventArgs
            {
                CreateBackup = createBackup,
                RememberChoice = RememberChoice
            });
        }
    }

    public class BackupConfirmationEventArgs : EventArgs
    {
        public bool CreateBackup { get; set; }
        public bool RememberChoice { get; set; }
    }
}