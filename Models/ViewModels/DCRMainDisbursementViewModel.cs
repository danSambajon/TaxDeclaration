namespace TaxDeclaration.Models.ViewModels
{
    public class DCRMainDisbursementViewModel
    {
        public string VoucherNo { get; set; }
        public string StnCode { get; set; } // USED
        public string AccountNo { get; set; } // USED
        public DateOnly? CashPoDate { get; set; } // USED
        public DateOnly? DcrDate { get; set; } // USED
        public string Company { get; set; } = "FILPRIDE";
    }
}
