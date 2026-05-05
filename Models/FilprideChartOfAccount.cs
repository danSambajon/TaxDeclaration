using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace TaxDeclaration.Models
{
    public class FilprideChartOfAccount
    {
        public int? AccountId { get; set; }
        public bool? IsMain { get; set; }
        public string? AccountNumber { get; set; }
        public string? AccountName { get; set; }
        public string? AccountType { get; set; }
        public string? NormalBalance { get; set; }
        public int? Level { get; set; }
        public string? CreatedBy { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string? EditedBy { get; set; }
        public DateTime? EditedDate { get; set; }
        public bool? HasChildren { get; set; }
        public int? ParentAccountId { get; set; }
        public string? FinancialStatementType { get; set; }
        public bool? IsHidden { get; set; }
    }
}
