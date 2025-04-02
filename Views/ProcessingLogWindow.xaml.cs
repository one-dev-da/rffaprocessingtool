using System.Windows;
using System.Windows.Controls;

namespace RffaDataComparisonTool.Views
{
    public partial class ProcessingLogWindow : Window
    {
        public ProcessingLogWindow()
        {
            InitializeComponent();

            // Prevent this window from minimizing the main application when closed with X
            this.Closing += (s, e) =>
            {
                // Set Owner to null before closing to prevent minimizing main window
                this.Owner = null;
            };
        }

        public void AppendLog(string message)
        {
            LogTextBox.AppendText(message + "\n");
            ScrollViewer.ScrollToEnd();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Set Owner to null before closing to prevent minimizing main window
            this.Owner = null;
            this.Close();
        }
    }
}