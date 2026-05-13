namespace TaxDeclaration.Models.ViewModels
{
    public class CvEntriesViewModel
    {

        public string? CvNo { get; set; } //

        public DateOnly? TranDate { get; set; } //

        public string? Payee { get; set; } //

        public string? BankCode { get; set; } //

        public string? BankName { get; set; } //

        public string? CheckNo { get; set; } //

        public DateOnly? CheckDate { get; set; } //

        public string? Particulars { get; set; } //

        public string? BsNo { get; set; } //

        public DateOnly? ChkClear { get; set; } //

        public string? Reference { get; set; } //

        public decimal? CvAmount { get; set; }

        public decimal? Amount { get; set; }

        public bool DrCr { get; set; }

        public string Acctcd {  get; set; }

        public string AcctName {  get; set; }

        public decimal? ColNum { get; set; }

        public string? Category { get; set; }

        public int? Seqid { get; set; }

        public string? CvType { get; set; }

        public bool IsDisplayEntry { get; set; }

        public DateOnly? DateFrom { get; set; }

        public DateOnly? DateTo { get; set; }

        public string? PayorName { get; set; }
        public string? PayeeName { get; set; }
        public string? Tin { get; set; }
        public string? AtcCode { get; set; }
        public string? Desc { get; set; }
        public decimal? Percent { get; set; }
        public decimal? NTotal { get; set; }
        public decimal? NAmount { get; set; }
        public decimal? NMonth { get; set; }
        public decimal? NYear { get; set; }
        public string? AccountNo { get; set; }
        public DateOnly? CashpoDate { get; set; }
        public DateOnly? DcrDate { get; set; }


    }
}
