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
            var sql = $@"
                SELECT check_voucher_header_id,
                check_voucher_header_no,
                date,
                cv_type,
                type,
                category,
                reference,
                particulars,
                total,
                check_no,
                check_date,
                bank_account_number,
                payee,
                created_by
                FROM public.""{tableName}""
                WHERE posted_by IS NOT NULL
                AND date >= @startDate
                AND date <= @endDate
                LIMIT 10";


            await using var cmd = new NpgsqlCommand(sql, conn);

            cmd.Parameters.AddWithValue("startDate", new DateOnly(2025, 12, 1));
            cmd.Parameters.AddWithValue("endDate", new DateOnly(2025, 12, 25));

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
                    Reference = reader.IsDBNull("reference") ? null : reader.GetString("reference"),
                    Particulars = reader.IsDBNull("particulars") ? null : reader.GetString("particulars"),
                    Total = reader.IsDBNull("total") ? null : reader.GetDecimal("total"),
                    CheckNo = reader.IsDBNull("check_no") ? null : reader.GetString("check_no"),
                    CheckDate = reader.IsDBNull("check_date") ? null : reader.GetFieldValue<DateOnly>("check_date"),
                    BankAccountNumber = reader.IsDBNull("bank_account_number") ? null : reader.GetString("bank_account_number"),
                    Payee = reader.IsDBNull("payee") ? null : reader.GetString("payee"),
                    CreatedBy = reader.IsDBNull("created_by") ? null : reader.GetString("created_by")
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
