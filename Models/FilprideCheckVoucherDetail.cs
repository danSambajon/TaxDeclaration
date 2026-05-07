using Microsoft.AspNetCore.Mvc.Rendering;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using TaxDeclaration.Models.ViewModels;

namespace TaxDeclaration.Models
{
    public class FilprideCheckVoucherDetail
    {
        public int? CheckVoucherDetailId { get; set; }
        public string? AccountNo { get; set; }
        public string? AccountName { get; set; }
        public string? TransactionNo { get; set; }
        public decimal? Debit { get; set; }
        public decimal? Credit { get; set; }
        public int? CheckVoucherHeaderId { get; set; }
        public decimal? Amount { get; set; }
        public decimal? AmountPaid { get; set; }
        public int? SupplierId { get; set; }
        public decimal? EwtPercent { get; set; }
        public bool IsUserSelected { get; set; }
        public bool IsVatable { get; set; }
        public int? BankId { get; set; }
        public int? CompanyId { get; set; }
        public int? CustomerId { get; set; }
        public int? EmployeeId { get; set; }
        public bool IsDisplayEntry { get; set; }
        public int? SubAccountType { get; set; }
        public int? SubAccountId { get; set; }
        public string? SubAccountName { get; set; }
        public FilprideCheckVoucherHeader? Header { get; set; }
        public TaxDeclareCompanyViewModel? CompanyVm { get; set; }
    }
}
