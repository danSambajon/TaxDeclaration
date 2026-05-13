using OfficeOpenXml;
using OfficeOpenXml.Style;
using TaxDeclaration.Models.ViewModels;

namespace TaxDeclaration.Services.Excel
{
    public class IBSDisbursementExcelService
    {
        public void ProcessReport1(ExcelWorksheet worksheet, GenerateTaxDeclarationViewModel viewModel)
        {

            #region == Title Area ==

            var mergedCells = worksheet.Cells["A1:C1"];
            mergedCells.Merge = true;
            mergedCells.Value = $"IBS DISBURSEMENT VOUCHERS - {viewModel.SelectedCompany}";
            mergedCells.Style.Font.Size = 13;
            mergedCells.Style.Font.Bold = true;

            worksheet.Cells[1, 5].Value = $"Report generated: {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}";

            mergedCells = worksheet.Cells["A2:C2"];
            mergedCells.Merge = true;
            mergedCells.Value = $"Voucher's Check Date from {viewModel.DateFrom} to {viewModel.DateTo}";

            mergedCells = worksheet.Cells["A3:C3"];
            mergedCells.Merge = true;
            mergedCells.Value = "Both Remitted and Unremitted";

            #endregion == Title Area ==

            #region == Headers ==

            var row = 4;
            var col = 1;

            var colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "VOUCHER #"; col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "VOUCHER DATE"; col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "PAYEE"; col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "PARTICULAR"; col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "CHECK #"; col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "CHECK DATE"; col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "BANK ACCT."; col++;
            var colSpandEnd = col - 1;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


            colSpanStart = col;
            worksheet.Cells[row, col].Value = "FROM DCR";
            worksheet.Cells[row + 1, col].Value = "DCR DATE"; col++;
            colSpandEnd = col - 1;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "VCH AMOUNT"; col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "BSNO"; col = col + 2;
            colSpandEnd = col - 2;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row, col+5]; mergedCells.Merge = true; mergedCells.Value = "TAX BASE PER CV ENTRY";
            mergedCells.Style.Fill.PatternType = ExcelFillStyle.Solid; mergedCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.GreenYellow);
            worksheet.Cells[row + 1, col].Value = "AMOUNT"; col++;
            worksheet.Cells[row + 1, col].Value = "ACCT#"; col++;
            worksheet.Cells[row + 1, col].Value = "ACCTNAME"; col++;
            worksheet.Cells[row + 1, col].Value = "INPUT VAT AMT"; col++;
            worksheet.Cells[row + 1, col].Value = "ACCT#"; col++;
            worksheet.Cells[row + 1, col].Value = "ACCTNAME"; col = col+2;
            colSpandEnd = col - 2;

            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row, col + 6]; mergedCells.Merge = true; mergedCells.Value = "FROM 2307 SYSTEM";
            mergedCells.Style.Fill.PatternType = ExcelFillStyle.Solid; mergedCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            worksheet.Cells[row + 1, col].Value = "PAYOR"; col++;
            worksheet.Cells[row + 1, col].Value = "PAYEE"; col++;
            worksheet.Cells[row + 1, col].Value = "TIN"; col++;
            worksheet.Cells[row + 1, col].Value = "PERCENT"; col++;
            worksheet.Cells[row + 1, col].Value = "EWT AMOUNT"; col++;
            worksheet.Cells[row + 1, col].Value = "MONTH"; col++;
            worksheet.Cells[row + 1, col].Value = "YEAR"; col = col+2;
            colSpandEnd = col - 2;

            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row, col + 1]; mergedCells.Merge = true; mergedCells.Value = "VAT";
            mergedCells.Style.Fill.PatternType = ExcelFillStyle.Solid; mergedCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            worksheet.Cells[row + 1, col].Value = "SHOULD BE TAX BASE"; col++;
            worksheet.Cells[row + 1, col].Value = "VARIANCE ON TAX BASE"; col = col+2;
            colSpandEnd = col - 2;

            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row, col + 2]; mergedCells.Merge = true; mergedCells.Value = "WITHHOLDING TAX";
            mergedCells.Style.Fill.PatternType = ExcelFillStyle.Solid; mergedCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            worksheet.Cells[row + 1, col].Value = "SHOULD BE TAX BASE"; col++;
            worksheet.Cells[row + 1, col].Value = "RATE"; col++;
            worksheet.Cells[row + 1, col].Value = "VARIANCE ON TAX BASE"; col = col+2;
            colSpandEnd = col - 2;

            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "COST (50)"; col++;
            colSpandEnd = col - 1;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

            
            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row, col + 1]; mergedCells.Merge = true; mergedCells.Value = "EXPENSE (55/65)";
            mergedCells.Style.Fill.PatternType = ExcelFillStyle.Solid; mergedCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            worksheet.Cells[row + 1, col].Value = "DEBIT"; col++;
            worksheet.Cells[row + 1, col].Value = "CREDIT"; col++;
            colSpandEnd = col - 1;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "CAPEX (102010)"; col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "VAT (101060200)"; col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "DEFERRED VAT (101060300)"; col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "EWT"; col = col+2;
            colSpandEnd = col - 2;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "UNCLEARED CHECKS"; col++;
            colSpandEnd = col - 1;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);








            #endregion == Headers ==

            // Populate the data rows
            var currencyFormat = "#,##0.00";

            // Auto-fit columns for better readability
            worksheet.Cells.AutoFitColumns();
            worksheet.View.FreezePanes(8, 1);
        }
    }
}
