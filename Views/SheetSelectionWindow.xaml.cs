using System.Windows;
using RffaDataComparisonTool.ViewModels;

namespace RffaDataComparisonTool.Views
{
    public partial class SheetSelectionWindow : Window
    {
        public SheetSelectionWindow()
        {
            InitializeComponent();

            // Close window when ViewModel's DialogResult is set
            DataContextChanged += (sender, args) =>
            {
                if (args.NewValue is SheetSelectionViewModel viewModel)
                {
                    viewModel.PropertyChanged += (s, e) =>
                    {
                        if (e.PropertyName == nameof(viewModel.DialogResult) && viewModel.DialogResult.HasValue)
                        {
                            DialogResult = viewModel.DialogResult;
                        }
                    };
                }
            };
        }
    }
}