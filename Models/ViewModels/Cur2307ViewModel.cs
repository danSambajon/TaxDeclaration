namespace TaxDeclaration.Models.ViewModels
{
    public class Cur2307ViewModel
    {
        public DateOnly? DateFrom { get; set; }
        public DateOnly? DateTo { get; set; }
        public string? CvNo { get; set; }
        public string? VoucherNo { get; set; }
        public string? AtcCode { get; set; }
        public decimal? Total { get; set; }
        public decimal? Amount { get; set; }
        public string? StationCode { get; set; }
        public bool Approved { get; set; }
        public int? RMonth { get; set; }
        public int? RYear { get; set; }
        public string? Desc { get; set; }
        public decimal? Percent { get; set; }
        public string? DownloadFrom { get; set; }
        public bool Cancelled { get; set; }
        public string? PayeeCode { get; set; }
        public string? PayeeName { get; set; }
        public string? TinA { get; set; }
        public string? TinB { get; set; }
        public string? TinC { get; set; }
        public string? TinD { get; set; }
        public string? PayorCode { get; set; }
        public string? PayorName { get; set; }
        public string? Description { get; set; }
        //a.DateFrom,
        //a.DateTo,
        //a.CvNo,
        //a.VoucherNo,
        //a.AtcCode,
        //a.Total,
        //a.Amount,
        //a.StationCode,
        //a.Approved,
        //a.RMonth,
        //a.RYear,
        //desc = (b == null || b.Desc == null) ? new string (' ', 254) : b.Desc,
        //percent = (b == null || b.Percent == null) ? 0 : b.Percent,
        //a.DownloadFrom,
        //a.Cancelled,
        //a.PayeeCode,
        //payeename = c == null ? null : c.PayeeName,
        //tin_a = c == null ? null : c.TinA,
        //tin_b = c == null ? null : c.TinB,
        //tin_c = c == null ? null : c.TinC,
        //tin_d = c == null ? null : c.TinD,
        //payorcode = d == null ? null : d.Code,
        //payorname = d == null ? null : d.Name,
        //a.Description
    }
}
