namespace TaxDeclaration.Models.ViewModels
{
    public class CvdetailxViewModel
    {
        public string? CvNo { get; set; }

        public int? SeqId { get; set; }

        public decimal? Amount { get; set; }

        public bool DrCr { get; set; }

        public string? Acctcd { get; set; }

        public bool IsDisplayEntry { get; set; }

        public FilprideCheckVoucherHeader? Header { get; set; }
    }
}
