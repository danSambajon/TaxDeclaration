using Microsoft.AspNetCore.Mvc.Rendering;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace TaxDeclaration.Models
{
    public class FilprideCheckVoucherDetail
    {
        public int? CheckVoucherDetailId { get; set; }
        public string? AccountNo { get; set; }
        public string? TransactionNo { get; set; }
        public decimal? Debit { get; set; }
        public decimal? Credit { get; set; }
        public decimal? Amount { get; set; }
        public bool? IsDisplayEntry { get; set; }
    }
}
