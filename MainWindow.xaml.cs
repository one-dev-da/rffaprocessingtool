using RffaDataComparisonTool.Services;
using RffaDataComparisonTool.ViewModels;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
namespace RffaDataComparisonTool
{
    public partial class MainWindow : Window
    {
        private readonly MainViewModel _viewModel;
        public MainWindow()
        {
            InitializeComponent();
            // Force layout update and measure
            this.Loaded += (s, e) =>
            {
                // Give the UI time to layout
                System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                    System.Windows.Threading.DispatcherPriority.Loaded,
                    new System.Action(() => {
                        // Force the scroll viewer to measure again
                        var scrollViewer = FindScrollViewer(this);
                        if (scrollViewer != null)
                        {
                            scrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Visible;
                            scrollViewer.UpdateLayout();
                            scrollViewer.ScrollToHome();
                        }
                    }));
            };

            // Setup services
            var historyService = new HistoryService();
            var userPreferencesService = new UserPreferencesService();
            var excelProcessor = new ExcelProcessorService(historyService, userPreferencesService);

            // Create loading service
            var loadingService = new LoadingService(
                showLoadingAction: message => ShowLoading(message),
                hideLoadingAction: () => HideLoading()
            );

            // Initialize ViewModel with all required services
            _viewModel = new MainViewModel(excelProcessor, historyService, loadingService, userPreferencesService);
            DataContext = _viewModel;

            // Configure DataGrid selection
            if (this.FindName("HistoryGrid") is DataGrid historyGrid)
            {
                historyGrid.SelectionChanged += (sender, e) =>
                {
                    // Store the selected items for use in the command binding
                    if (historyGrid.SelectedItems.Count > 0)
                    {
                        var selectedItems = new List<Models.ProcessingRecord>();
                        foreach (var item in historyGrid.SelectedItems)
                        {
                            if (item is Models.ProcessingRecord record)
                            {
                                selectedItems.Add(record);
                            }
                        }
                        _viewModel.UpdateSelectedHistoryItems(selectedItems);
                    }
                };
            }
        }

        private ScrollViewer FindScrollViewer(DependencyObject parent)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is ScrollViewer scrollViewer)
                    return scrollViewer;
                var result = FindScrollViewer(child);
                if (result != null)
                    return result;
            }
            return null;
        }

        public void ShowLoading(string message)
        {
            // Update UI on the UI thread and force refresh
            Application.Current.Dispatcher.Invoke(() =>
            {
                LoadingMessage.Text = message;
                LoadingOverlay.Visibility = Visibility.Visible;

                // Force UI update
                UpdateLayout();
            });
        }

        public void HideLoading()
        {
            // Update UI on the UI thread
            Application.Current.Dispatcher.Invoke(() =>
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
            });
        }
    }
}