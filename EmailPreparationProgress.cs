namespace PersonalizedEmailSender;

internal sealed record EmailPreparationProgress(
    string Status,
    int Current,
    int Total);
