using System.Windows;

namespace PersonalizedEmailSender;

public partial class EmailPreparationProgressWindow : Window
{
    private double _lastRatio;

    public EmailPreparationProgressWindow()
    {
        InitializeComponent();
    }

    public void Report(string status, int current, int total)
    {
        Dispatcher.Invoke(() =>
        {
            int safeTotal = Math.Max(total, 1);
            int safeCurrent = Math.Clamp(current, 0, safeTotal);
            double ratio = safeCurrent / (double)safeTotal;
            _lastRatio = ratio;

            StatusTextBlock.Text = status;
            ProgressFillBorder.Width = ProgressTrackGrid.ActualWidth * ratio;
            ProgressPercentTextBlock.Text = $"{ratio:P0}";
            ProgressBarTextBlock.Text = BuildTextProgressBar(ratio);
        });
    }

    private void ProgressTrackGrid_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        ProgressFillBorder.Width = ProgressTrackGrid.ActualWidth * _lastRatio;
    }

    private static string BuildTextProgressBar(double ratio)
    {
        const int totalSegments = 24;
        int filledSegments = (int)Math.Round(totalSegments * ratio);
        return $"{new string('|', filledSegments)}{new string('-', totalSegments - filledSegments)}";
    }
}
