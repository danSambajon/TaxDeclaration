using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace TaxDeclaration.Models
{
    public class FilprideCheckVoucherHeader
    {
        // Primary Key
        public int CheckVoucherHeaderId { get; set; }

        // Voucher Info
        public string? CheckVoucherHeaderNo { get; set; }
        public DateOnly? Date { get; set; }
        public string? CvType { get; set; }
        public string? Type { get; set; }
        public string? Category { get; set; }
        public string? Reference { get; set; }
        public string? Particulars { get; set; }

        // Financial
        public decimal? Total { get; set; }
        public string? CheckNo { get; set; }
        public DateOnly? CheckDate { get; set; }
        public string? BankAccountNumber { get; set; }
        public string? Payee { get; set; }

        // Audit Trail
        public string? CreatedBy { get; set; }
    }
}
