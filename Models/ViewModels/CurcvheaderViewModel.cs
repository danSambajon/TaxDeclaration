namespace TaxDeclaration.Models.ViewModels
{
    public class CurcvheaderViewModel
    {
        public string? CvDtl { get; set; }

        public CvheaderxViewModel? Header { get; set; }

        public DateOnly? CashpoDate { get; set; }

        public DateOnly? DCRDate { get; set; }

        public string? AccountNo { get; set; }

        public string? StnCode { get; set; }

        public string? Company { get; set; }

        public string? Rem { get; set; }

        public DateOnly? DateFrom { get; set; }

        public DateOnly? DateTo { get; set; }

        public string? AtcCode { get; set; }

        public string? AtcDesc { get; set; }

        public decimal Percent { get; set; }

        public string? PayorName { get; set; }

        public string? PayeeName { get; set; }

        public string? Tin { get; set; }

        public decimal Total { get; set; }

        public decimal Amt2307 { get; set; }

        public decimal NMonth { get; set; }

        public decimal NYear { get; set; }

        public decimal VatAmt { get; set; }

        public string? VatAcctNo { get; set; }

        public string? VatDesc { get; set; }

        public decimal EwtAmt { get; set; }

        public string? EwtAcctNo { get; set; }

        public string? EwtDesc { get; set; }

        public decimal VatShouldBe { get; set; }

        public decimal VatVariance { get; set; }

        public bool? VatOk { get; set; }

        public decimal EwtShouldBe { get; set; }

        public decimal EwtVariance { get; set; }

        public bool? EwtOk { get; set; }

        public decimal LedgerDebit { get; set; }

        public string? LedgerAcctNo { get; set; }

        public string? LedgerDesc { get; set; }

        public string? RowNum { get; set; }

        public string? CvType { get; set; }
    }
}
