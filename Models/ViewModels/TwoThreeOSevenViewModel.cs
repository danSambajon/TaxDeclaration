namespace TaxDeclaration.Models.ViewModels
{
    public class TwoThreeOSevenViewModel
    {
        public int RecId { get; set; }
        public string? ControlNum { get; set; }
        public DateTime? DateFrom { get; set; } // 
        public DateTime? DateTo { get; set; } //
        public string? VoucherNo { get; set; } //
        public string? CvNo { get; set; } //
        public string? PayeeCode { get; set; } //
        public string? AtcCode { get; set; } //
        public decimal? Total { get; set; } //
        public decimal? Amount { get; set; } //
        public string? DownloadFrom { get; set; } //
        public string? StationCode { get; set; } //
        public string? Description { get; set; } //
        public bool Approved { get; set; } //
        public int? RMonth { get; set; } //
        public int? RYear { get; set; } //
        public bool Cancelled { get; set; } //
    }
}
