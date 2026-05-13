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
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "VOUCHER #"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "VOUCHER DATE"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "PAYEE"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "PARTICULAR"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "CHECK #"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "CHECK DATE"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "BANK ACCT."; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            var colSpandEnd = col - 1;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


            colSpanStart = col;
            worksheet.Cells[row, col].Value = "FROM DCR"; worksheet.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[row + 1, col].Value = "DCR DATE"; worksheet.Cells[row+1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            colSpandEnd = col - 1;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "VCH AMOUNT"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "BSNO"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col = col + 2;
            colSpandEnd = col - 2;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row, col+5]; mergedCells.Merge = true; mergedCells.Value = "TAX BASE PER CV ENTRY"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            mergedCells.Style.Fill.PatternType = ExcelFillStyle.Solid; mergedCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.GreenYellow);
            worksheet.Cells[row + 1, col].Value = "AMOUNT"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "ACCT#"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "ACCTNAME"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "INPUT VAT AMT"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "ACCT#"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "ACCTNAME"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col = col+2;
            colSpandEnd = col - 2;

            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row, col + 6]; mergedCells.Merge = true; mergedCells.Value = "FROM 2307 SYSTEM"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            mergedCells.Style.Fill.PatternType = ExcelFillStyle.Solid; mergedCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            worksheet.Cells[row + 1, col].Value = "PAYOR"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "PAYEE"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "TIN"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "PERCENT"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "EWT AMOUNT"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "MONTH"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "YEAR"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col = col+2;
            colSpandEnd = col - 2;

            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row, col + 1]; mergedCells.Merge = true; mergedCells.Value = "VAT"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            mergedCells.Style.Fill.PatternType = ExcelFillStyle.Solid; mergedCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            worksheet.Cells[row + 1, col].Value = "SHOULD BE TAX BASE"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "VARIANCE ON TAX BASE"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col = col+2;
            colSpandEnd = col - 2;

            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row, col + 2]; mergedCells.Merge = true; mergedCells.Value = "WITHHOLDING TAX"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            mergedCells.Style.Fill.PatternType = ExcelFillStyle.Solid; mergedCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            worksheet.Cells[row + 1, col].Value = "SHOULD BE TAX BASE"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "RATE"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "VARIANCE ON TAX BASE"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col = col+2;
            colSpandEnd = col - 2;

            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "COST (50)"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            colSpandEnd = col - 1;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

            
            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row, col + 1]; mergedCells.Merge = true; mergedCells.Value = "EXPENSE (55/65)"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            mergedCells.Style.Fill.PatternType = ExcelFillStyle.Solid; mergedCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            worksheet.Cells[row + 1, col].Value = "DEBIT"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "CREDIT"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            colSpandEnd = col - 1;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "CAPEX (102010)"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "VAT (101060200)"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "DEFERRED VAT (101060300)"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "EWT"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col = col+2;
            colSpandEnd = col - 2;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


            colSpanStart = col;
            mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "UNCLEARED CHECKS"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            colSpandEnd = col - 1;

            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

            worksheet.Cells[row, 1, row + 1, 42].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, 1, row + 1, 42].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[row, 1, row + 1, 42].Style.Font.Bold = true;

            #endregion == Headers ==

            #region == Column ==

            var currencyFormat = "#,##0.00";

            worksheet.View.FreezePanes(6, 1);
            worksheet.Columns.AutoFit();

            for (int ctr = 1; ctr != 45; ctr++)
            {
                worksheet.Column(ctr).Width = 20;
            }

            foreach (var colTemp in new List<int> { 11, 18, 26, 29, 33, 41 })
            {
                worksheet.Column(colTemp).Width = 1;
            }

            #endregion == Column ==
        }
    }
}
