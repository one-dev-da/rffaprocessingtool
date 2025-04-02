using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace RffaDataComparisonTool.Views
{
    public partial class LoadingOverlay : UserControl
    {
        public LoadingOverlay()
        {
            InitializeComponent();

            // Start the spinner animation
            DoubleAnimation rotateAnimation = new DoubleAnimation
            {
                From = 0,
                To = 360,
                Duration = TimeSpan.FromSeconds(1.5),
                RepeatBehavior = RepeatBehavior.Forever
            };

            SpinnerRotation.BeginAnimation(RotateTransform.AngleProperty, rotateAnimation);
        }

        // Message property
        public string Message
        {
            get { return MessageText.Text; }
            set { MessageText.Text = value; }
        }
    }
}