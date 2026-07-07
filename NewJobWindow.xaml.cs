using System.Windows;

namespace PersonalizedEmailSender;

public partial class NewJobWindow : Window
{
    public PersonalizedEmailJob? CreatedJob { get; private set; }

    public NewJobWindow()
    {
        InitializeComponent();
        NewJobScreen.StartNewJobSetup();
    }

    private void NewJobScreen_BackRequested(object sender, EventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void NewJobScreen_JobCreated(object sender, PersonalizedEmailJob job)
    {
        EmailContentWindow contentWindow = new(job)
        {
            Owner = this
        };

        contentWindow.ShowDialog();

        if (!contentWindow.HasSavedDraft)
        {
            StatusTextBlock.Text = "Email content editing was cancelled.";
            return;
        }

        CreatedJob = contentWindow.CurrentJob;
        StatusTextBlock.Text = $"Draft saved to {DraftStore.DraftsFolderPath}.";
        DialogResult = true;
        Close();
    }

    private void NewJobScreen_StatusChanged(object sender, string status)
    {
        StatusTextBlock.Text = status;
    }
}
