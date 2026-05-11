namespace TaxDeclaration.Models.ViewModels
{
    public class TwoThreeOSevenViewModel
    {
        public int RecId { get; set; }
        public string? ControlNum { get; set; }
        public DateOnly? DateFrom { get; set; } // 
        public DateOnly? DateTo { get; set; } //
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
        public string? CompanyCode { get; set; }
    }

    public class ATCViewModel
    {
        public string? Desc {  get; set; }
        public decimal? Percent {  get; set; }
        public string? Code {  get; set; }
    }

    public class PayeeViewModel
    {
        public string? PayeeName {  get; set; }
        public string? PayeeCode {  get; set; }
        public string? TinA {  get; set; }
        public string? TinB {  get; set; }
        public string? TinC {  get; set; }
        public string? TinD {  get; set; }
    }

    public class EwtCompanyViewModel
    {
        public string? Code { get; set; }
        public string? Name { get; set; }
    }
}
