namespace TaxDeclaration.Models.ViewModels
{
    public class DCRMainDisbursementViewModel
    {
        public string StnCode { get; set; } // USED
        public string AccountNo { get; set; } // USED
        public DateTime? CashPoDate { get; set; } // USED
        public DateTime? DcrDate { get; set; } // USED
        public string Company { get; set; } = "FILPRIDE";
    }
}
