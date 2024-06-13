namespace IntugentWebApp.Models
{
    public class MFGSearchResult
    {
        public int ID { get; set; }
        public DateTime? Date { get; set; }
        public string? Shift { get; set; }
        public DateTime? GreenBoardTimeStamp { get; set; }
        public DateTime? FGBoardTimeStamp { get; set; }
        public string? Product { get; set; }
        public string? Location { get; set; }
    }
}
