using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Npgsql;
using System.Collections.Generic;
using System.Data;
using TaxDeclaration.Models;

namespace TaxDeclaration.Services
{
    public class IBSConnectionService
    {
        private readonly IConfiguration _configuration;

        public IBSConnectionService(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public async Task<List<FilprideCheckVoucherHeader>> GetCheckVoucherHeaders ()
        {
            var listOfEntries = new List<FilprideCheckVoucherHeader> ();
            var connString = _configuration.GetConnectionString("IBSConnection");
            var tableName = "filpride_check_voucher_headers";

            await using var conn = new NpgsqlConnection(connString);

            await conn.OpenAsync();

            // Row count
            var countSql = $"SELECT COUNT(*) FROM public.\"{tableName}\"";
            await using var countCmd = new NpgsqlCommand(countSql, conn);
            var count = Convert.ToInt32(await countCmd.ExecuteScalarAsync());

            // Fields + sample data
            var sql = $"SELECT * FROM public.\"{tableName}\" LIMIT 10";
            await using var cmd = new NpgsqlCommand(sql, conn);
            await using var reader = await cmd.ExecuteReaderAsync();
            var schema = reader.GetColumnSchema();

            while (await reader.ReadAsync())
            {
                var header = new FilprideCheckVoucherHeader
                {
                    CheckVoucherHeaderId = reader.GetInt32("check_voucher_header_id"),
                    CheckVoucherHeaderNo = reader.IsDBNull("check_voucher_header_no") ? null : reader.GetString("check_voucher_header_no"),
                    Date = reader.IsDBNull("date") ? null : reader.GetFieldValue<DateOnly>("date"),
                    CvType = reader.IsDBNull("cv_type") ? null : reader.GetString("cv_type"),
                    Type = reader.IsDBNull("type") ? null : reader.GetString("type"),
                    Category = reader.IsDBNull("category") ? null : reader.GetString("category"),
                    Status = reader.IsDBNull("status") ? null : reader.GetString("status"),
                    Reference = reader.IsDBNull("reference") ? null : reader.GetString("reference"),
                    Particulars = reader.IsDBNull("particulars") ? null : reader.GetString("particulars"),
                    CancellationRemarks = reader.IsDBNull("cancellation_remarks") ? null : reader.GetString("cancellation_remarks"),
                    OldCvNo = reader.IsDBNull("old_cv_no") ? null : reader.GetString("old_cv_no"),
                    Company = reader.IsDBNull("company") ? null : reader.GetString("company"),
                    RrNo = reader.IsDBNull("rr_no") ? null : reader.GetFieldValue<string[]>("rr_no"),
                    SiNo = reader.IsDBNull("si_no") ? null : reader.GetFieldValue<string[]>("si_no"),
                    PoNo = reader.IsDBNull("po_no") ? null : reader.GetFieldValue<string[]>("po_no"),
                    Amount = reader.IsDBNull("amount") ? null : reader.GetFieldValue<decimal[]>("amount"),
                    Total = reader.IsDBNull("total") ? null : reader.GetDecimal("total"),
                    CheckAmount = reader.IsDBNull("check_amount") ? null : reader.GetDecimal("check_amount"),
                    AmountPaid = reader.IsDBNull("amount_paid") ? null : reader.GetDecimal("amount_paid"),
                    InvoiceAmount = reader.IsDBNull("invoice_amount") ? null : reader.GetDecimal("invoice_amount"),
                    TaxPercent = reader.IsDBNull("tax_percent") ? null : reader.GetDecimal("tax_percent"),
                    TaxType = reader.IsDBNull("tax_type") ? null : reader.GetString("tax_type"),
                    VatType = reader.IsDBNull("vat_type") ? null : reader.GetString("vat_type"),
                    BankId = reader.IsDBNull("bank_id") ? null : reader.GetInt32("bank_id"),
                    CheckNo = reader.IsDBNull("check_no") ? null : reader.GetString("check_no"),
                    CheckDate = reader.IsDBNull("check_date") ? null : reader.GetFieldValue<DateOnly>("check_date"),
                    BankAccountName = reader.IsDBNull("bank_account_name") ? null : reader.GetString("bank_account_name"),
                    BankAccountNumber = reader.IsDBNull("bank_account_number") ? null : reader.GetString("bank_account_number"),
                    SupplierId = reader.IsDBNull("supplier_id") ? null : reader.GetInt32("supplier_id"),
                    SupplierName = reader.IsDBNull("supplier_name") ? null : reader.GetString("supplier_name"),
                    Payee = reader.IsDBNull("payee") ? null : reader.GetString("payee"),
                    Address = reader.IsDBNull("address") ? null : reader.GetString("address"),
                    Tin = reader.IsDBNull("tin") ? null : reader.GetString("tin"),
                    EmployeeId = reader.IsDBNull("employee_id") ? null : reader.GetInt32("employee_id"),
                    SupportingFileSavedFileName = reader.IsDBNull("supporting_file_saved_file_name") ? null : reader.GetString("supporting_file_saved_file_name"),
                    SupportingFileSavedUrl = reader.IsDBNull("supporting_file_saved_url") ? null : reader.GetString("supporting_file_saved_url"),
                    DcpDate = reader.IsDBNull("dcp_date") ? null : reader.GetFieldValue<DateOnly>("dcp_date"),
                    DcrDate = reader.IsDBNull("dcr_date") ? null : reader.GetFieldValue<DateOnly>("dcr_date"),
                    LiquidationDate = reader.IsDBNull("liquidation_date") ? null : reader.GetFieldValue<DateOnly>("liquidation_date"),
                    IsPaid = reader.GetBoolean("is_paid"),
                    IsPrinted = reader.GetBoolean("is_printed"),
                    IsAdvances = reader.GetBoolean("is_advances"),
                    IsPayroll = reader.GetBoolean("is_payroll"),
                    CreatedBy = reader.IsDBNull("created_by") ? null : reader.GetString("created_by"),
                    CreatedDate = reader.IsDBNull("created_date") ? null : reader.GetDateTime("created_date"),
                    EditedBy = reader.IsDBNull("edited_by") ? null : reader.GetString("edited_by"),
                    EditedDate = reader.IsDBNull("edited_date") ? null : reader.GetDateTime("edited_date"),
                    CanceledBy = reader.IsDBNull("canceled_by") ? null : reader.GetString("canceled_by"),
                    CanceledDate = reader.IsDBNull("canceled_date") ? null : reader.GetDateTime("canceled_date"),
                    VoidedBy = reader.IsDBNull("voided_by") ? null : reader.GetString("voided_by"),
                    VoidedDate = reader.IsDBNull("voided_date") ? null : reader.GetDateTime("voided_date"),
                    PostedBy = reader.IsDBNull("posted_by") ? null : reader.GetString("posted_by"),
                    PostedDate = reader.IsDBNull("posted_date") ? null : reader.GetDateTime("posted_date"),
                    ApprovedBy = reader.IsDBNull("approved_by") ? null : reader.GetString("approved_by"),
                    ApprovedDate = reader.IsDBNull("approved_date") ? null : reader.GetDateTime("approved_date"),
                };

                listOfEntries.Add(header);
            }

            return listOfEntries;
        }

        public async Task<List<FilprideCheckVoucherDetail>> GetCheckVoucherDetails ()
        {
            var listOfEntries = new List<FilprideCheckVoucherDetail> ();
            var connString = _configuration.GetConnectionString("IBSConnection");
            var tableName = "filpride_check_voucher_details";

            await using var conn = new NpgsqlConnection(connString);

            await conn.OpenAsync();

            // Row count
            var countSql = $"SELECT COUNT(*) FROM public.\"{tableName}\"";
            await using var countCmd = new NpgsqlCommand(countSql, conn);
            var count = Convert.ToInt32(await countCmd.ExecuteScalarAsync());

            // Fields + sample data
            var sql = $"SELECT * FROM public.\"{tableName}\" LIMIT 10";
            await using var cmd = new NpgsqlCommand(sql, conn);
            await using var reader = await cmd.ExecuteReaderAsync();
            var schema = reader.GetColumnSchema();

            while (await reader.ReadAsync())
            {
                var detail = new FilprideCheckVoucherDetail
                {
                    CheckVoucherDetailId = reader.IsDBNull("check_voucher_detail_id") ? null : reader.GetInt32("check_voucher_detail_id"),
                    AccountNo = reader.IsDBNull("account_no") ? null : reader.GetString("account_no"),
                    AccountName = reader.IsDBNull("account_name") ? null : reader.GetString("account_name"),
                    TransactionNo = reader.IsDBNull("transaction_no") ? null : reader.GetString("transaction_no"),
                    Debit = reader.IsDBNull("debit") ? null : reader.GetDecimal("debit"),
                    Credit = reader.IsDBNull("credit") ? null : reader.GetDecimal("credit"),
                    CheckVoucherHeaderId = reader.IsDBNull("check_voucher_header_id") ? null : reader.GetInt32("check_voucher_header_id"),
                    Amount = reader.IsDBNull("amount") ? null : reader.GetDecimal("amount"),
                    AmountPaid = reader.IsDBNull("amount_paid") ? null : reader.GetDecimal("amount_paid"),
                    SupplierId = reader.IsDBNull("supplier_id") ? null : reader.GetInt32("supplier_id"),
                    EwtPercent = reader.IsDBNull("ewt_percent") ? null : reader.GetDecimal("ewt_percent"),
                    IsUserSelected = reader.IsDBNull("is_user_selected") ? null : reader.GetBoolean("is_user_selected"),
                    IsVatable = reader.IsDBNull("is_vatable") ? null : reader.GetBoolean("is_vatable"),
                    BankId = reader.IsDBNull("bank_id") ? null : reader.GetInt32("bank_id"),
                    CompanyId = reader.IsDBNull("company_id") ? null : reader.GetInt32("company_id"),
                    CustomerId = reader.IsDBNull("customer_id") ? null : reader.GetInt32("customer_id"),
                    EmployeeId = reader.IsDBNull("employee_id") ? null : reader.GetInt32("employee_id"),
                    IsDisplayEntry = reader.IsDBNull("is_display_entry") ? null : reader.GetBoolean("is_display_entry"),
                    SubAccountType = reader.IsDBNull("sub_account_type") ? null : reader.GetInt32("sub_account_type"),
                    SubAccountId = reader.IsDBNull("sub_account_id") ? null : reader.GetInt32("sub_account_id"),
                    SubAccountName = reader.IsDBNull("sub_account_name") ? null : reader.GetString("sub_account_name"),
                };

                listOfEntries.Add(detail);
            }

            return listOfEntries;
        }

        public async Task<List<FilprideChartOfAccount>> GetChartOfAccounts ()
        {
            var listOfEntries = new List<FilprideChartOfAccount> ();
            var connString = _configuration.GetConnectionString("IBSConnection");
            var tableName = "filpride_chart_of_accounts";

            await using var conn = new NpgsqlConnection(connString);

            await conn.OpenAsync();

            // Row count
            var countSql = $"SELECT COUNT(*) FROM public.\"{tableName}\"";
            await using var countCmd = new NpgsqlCommand(countSql, conn);
            var count = Convert.ToInt32(await countCmd.ExecuteScalarAsync());

            // Fields + sample data
            var sql = $"SELECT * FROM public.\"{tableName}\" LIMIT 10";
            await using var cmd = new NpgsqlCommand(sql, conn);
            await using var reader = await cmd.ExecuteReaderAsync();
            var schema = reader.GetColumnSchema();

            while (await reader.ReadAsync())
            {
                var coa = new FilprideChartOfAccount
                {
                    AccountId = reader.IsDBNull("account_id") ? null : reader.GetInt32("account_id"),
                    IsMain = reader.IsDBNull("is_main") ? null : reader.GetBoolean("is_main"),
                    AccountNumber = reader.IsDBNull("account_number") ? null : reader.GetString("account_number"),
                    AccountName = reader.IsDBNull("account_name") ? null : reader.GetString("account_name"),
                    AccountType = reader.IsDBNull("account_type") ? null : reader.GetString("account_type"),
                    NormalBalance = reader.IsDBNull("normal_balance") ? null : reader.GetString("normal_balance"),
                    Level = reader.IsDBNull("level") ? null : reader.GetInt32("level"),
                    CreatedBy = reader.IsDBNull("created_by") ? null : reader.GetString("created_by"),
                    CreatedDate = reader.IsDBNull("created_date") ? null : reader.GetDateTime("created_date"),
                    EditedBy = reader.IsDBNull("edited_by") ? null : reader.GetString("edited_by"),
                    EditedDate = reader.IsDBNull("edited_date") ? null : reader.GetDateTime("edited_date"),
                    HasChildren = reader.IsDBNull("has_children") ? null : reader.GetBoolean("has_children"),
                    ParentAccountId = reader.IsDBNull("parent_account_id") ? null : reader.GetInt32("parent_account_id"),
                    FinancialStatementType = reader.IsDBNull("financial_statement_type") ? null : reader.GetString("financial_statement_type"),
                    IsHidden = reader.IsDBNull("is_hidden") ? null : reader.GetBoolean("is_hidden"),
                };

                listOfEntries.Add(coa);
            }

            return listOfEntries;
        }
    }
}
