namespace TaxDeclaration.Models.ViewModels
{
    public class CvheaderxViewModel
    {
        public string? CvNo { get; set; }

        public DateOnly? TranDate { get; set; }

        public string? Payee { get; set; }

        public string? BankCode { get; set; }

        public string? BankName { get; set; }

        public string? CheckNo { get; set; }

        public decimal? Amount { get; set; }

        public DateOnly? CheckDate { get; set; }

        public string? Particulars { get; set; }

        public string? CreatedBy { get; set; }

        public string? AcctCd { get; set; }

        public string? BsNo { get; set; }

        public DateOnly? CheckClearing { get; set; }

        public string? NameCategory { get; set; }

        public bool? IsCancelled { get; set; }

        public DateOnly? ChkClear { get; set; }

        public string? Type { get; set; }

        public string? Reference { get; set; }

        public string? Category { get; set; }

        public string? CvType { get; set; }
    }
}
