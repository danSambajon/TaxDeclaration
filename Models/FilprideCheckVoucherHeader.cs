using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace TaxDeclaration.Models
{
    public class FilprideCheckVoucherHeader
    {
        public int? CheckVoucherHeaderId { get; set; }
        public string? CheckVoucherHeaderNo { get; set; }
        public DateOnly? Date { get; set; }
        public string? CvType { get; set; }
        public string? Type { get; set; }
        public string? Category { get; set; }
        public string? Status { get; set; }
        public string? Reference { get; set; }
        public string? Particulars { get; set; }
        public string? CancellationRemarks { get; set; }
        public string? OldCvNo { get; set; }
        public string? Company { get; set; }
        public string[]? RrNo { get; set; }
        public string[]? SiNo { get; set; }
        public string[]? PoNo { get; set; }
        public decimal[]? Amount { get; set; }
        public decimal? Total { get; set; }
        public decimal? CheckAmount { get; set; }
        public decimal? AmountPaid { get; set; }
        public decimal? InvoiceAmount { get; set; }
        public decimal? TaxPercent { get; set; }
        public string? TaxType { get; set; }
        public string? VatType { get; set; }
        public int? BankId { get; set; }
        public string? CheckNo { get; set; }
        public DateOnly? CheckDate { get; set; }
        public string? BankAccountName { get; set; }
        public string? BankAccountNumber { get; set; }
        public int? SupplierId { get; set; }
        public string? SupplierName { get; set; }
        public string? Payee { get; set; }
        public string? Address { get; set; }
        public string? Tin { get; set; }
        public int? EmployeeId { get; set; }
        public string? SupportingFileSavedFileName { get; set; }
        public string? SupportingFileSavedUrl { get; set; }
        public DateOnly? DcpDate { get; set; }
        public DateOnly? DcrDate { get; set; }
        public DateOnly? LiquidationDate { get; set; }
        public bool? IsPaid { get; set; }
        public bool? IsPrinted { get; set; }
        public bool? IsAdvances { get; set; }
        public bool? IsPayroll { get; set; }
        public string? CreatedBy { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string? EditedBy { get; set; }
        public DateTime? EditedDate { get; set; }
        public string? CanceledBy { get; set; }
        public DateTime? CanceledDate { get; set; }
        public string? VoidedBy { get; set; }
        public DateTime? VoidedDate { get; set; }
        public string? PostedBy { get; set; }
        public DateTime? PostedDate { get; set; }
        public string? ApprovedBy { get; set; }
        public DateTime? ApprovedDate { get; set; }

        public List<FilprideCheckVoucherDetail> Details { get; set; } = new List<FilprideCheckVoucherDetail>();
        public List<FilprideCheckVoucherDetail> Referenced { get; set; } = new List<FilprideCheckVoucherDetail>();
    }
}
