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

        public async Task<List<FilprideCheckVoucherHeader>> GetCheckVoucherHeaders (DateOnly dateFrom, DateOnly dateTo)
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
                AND date <= @endDate";

            await using var cmd = new NpgsqlCommand(sql, conn);

            cmd.Parameters.AddWithValue("startDate", dateFrom);
            cmd.Parameters.AddWithValue("endDate", dateTo);

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

        public async Task<List<FilprideCheckVoucherDetail>> GetCheckVoucherDetails (List<string?>? cvHeaderTransNos)
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

            var transNos = cvHeaderTransNos?
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToArray();

            // Fields + sample data
            var sql = $@"
                SELECT check_voucher_detail_id,
                account_no,
                transaction_no, 
                debit,
                credit,
                amount,
                is_display_entry
                FROM public.""{tableName}""
                WHERE transaction_no = ANY(@transNos)";

            await using var cmd = new NpgsqlCommand(sql, conn);

            cmd.Parameters.AddWithValue("transNos", transNos ?? Array.Empty<string>());

            await using var reader = await cmd.ExecuteReaderAsync();
            var schema = reader.GetColumnSchema();

            while (await reader.ReadAsync())
            {
                var detail = new FilprideCheckVoucherDetail
                {
                    CheckVoucherDetailId = reader.IsDBNull("check_voucher_detail_id") ? null : reader.GetInt32("check_voucher_detail_id"),
                    AccountNo = reader.IsDBNull("account_no") ? null : reader.GetString("account_no"),
                    TransactionNo = reader.IsDBNull("transaction_no") ? null : reader.GetString("transaction_no"),
                    Debit = reader.IsDBNull("debit") ? null : reader.GetDecimal("debit"),
                    Credit = reader.IsDBNull("credit") ? null : reader.GetDecimal("credit"),
                    Amount = reader.IsDBNull("amount") ? null : reader.GetDecimal("amount"),
                    IsDisplayEntry = reader.IsDBNull("is_display_entry") ? null : reader.GetBoolean("is_display_entry"),
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
            var sql = $"SELECT * FROM public.\"{tableName}\"";
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
