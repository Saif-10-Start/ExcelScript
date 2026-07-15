namespace ExcelScript
{
    public class ProgressCallback
    {
        public Action<TimeSpan>? WorkbookLoaded { get; init; }
        public Action<TimeSpan>? DataCopied { get; init; }
        public Action<TimeSpan>? ItemsGrouped { get; init; }
        public Action<TimeSpan>? DataFormatted { get; init; }
        public Action<TimeSpan>? SummaryGenerated { get; init; }
        public Action<TimeSpan>? WorkbookStyled { get; init; }
        public Action<TimeSpan>? CommentsCopied { get; init; }
        public Action<TimeSpan>? WorkbookSaved { get; init; }
    }
}
