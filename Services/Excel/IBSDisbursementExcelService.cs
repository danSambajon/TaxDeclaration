using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Hosting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Composition;
using TaxDeclaration.Models.ViewModels;

namespace TaxDeclaration.Services.Excel
{
    public class IBSDisbursementExcelService
    {
        public void ProcessHeaderReport(ExcelWorksheet worksheet, 
            GenerateTaxDeclarationViewModel viewModel, 
            List<CurcvheaderViewModel> curcvheader,
            List<CvEntriesViewModel> cventries,
            bool isReport3)
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

            // var currencyFormat = "#,##0.00;[Red](#,##0.00)";
            var currencyFormat = "#,##0.00;[Red](#,##0.00);??;@";

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

            if (isReport3 == false)
            {
                colSpanStart = col;
                mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "UNCLEARED CHECKS"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
                colSpandEnd = col - 1;

                worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            }

            worksheet.Cells[row, 1, row + 1, 42].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, 1, row + 1, 42].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[row, 1, row + 1, 42].Style.Font.Bold = true;

            #endregion == Headers ==

            #region == Values ==

            var vchAmountTotal = 0m;
            var perCvEntryAmountTotal = 0m;
            var perCvEntryInputVatAmountTotal = 0m;
            var vatShouldBeTaxTotal = 0m;
            var vatVarianceonTaxTotal = 0m;
            var whtVarianceonTaxTotal = 0m;
            var cost50Total = 0m;
            var expenseDebitTotal = 0m;
            var expenseCreditTotal = 0m;
            var expenseCapexTotal = 0m;
            var expenseVatTotal = 0m;
            var expenseDVatTotal = 0m;
            var expenseEwtTotal = 0m;
            var unclearedChecksTotal = 0m;

            row = 6;

            foreach (var cv in curcvheader)
            {
                col = 1;

                worksheet.Cells[row, col].Value = cv.Header!.Reference != null ? $"{cv.Header.CvNo} / {cv.Header.Reference}" : cv.Header.CvNo; col++;

                worksheet.Cells[row, col].Value = cv.Header.TranDate; col++; // 2
                worksheet.Cells[row, col].Value = cv.Header.Payee?.Trim(); col++; //3
                worksheet.Cells[row, col].Value = cv.Header.Particulars?.Trim(); col++; //4
                worksheet.Cells[row, col].Value = cv.Header.CheckNo?.Trim(); col++; //5
                worksheet.Cells[row, col].Value = cv.Header.CheckDate.ToString()?.Trim(); col++; // 6 date
                worksheet.Cells[row, col].Value = cv.Header.BankCode; col++; //7
                worksheet.Cells[row, col].Value = cv.DCRDate; col++; //8
                worksheet.Cells[row, col].Value = cv.Header.Amount; col++; vchAmountTotal += cv.Header.Amount ?? 0m; // 9
                worksheet.Cells[row, col].Value = cv.Header.BsNo; col++; // 10
                col++; // 11
                worksheet.Cells[row, col].Value = cv.VatAmt; col++; perCvEntryAmountTotal += cv.VatAmt; // 12
                worksheet.Cells[row, col].Value = cv.VatAcctNo; col++; // 13
                worksheet.Cells[row, col].Value = cv.VatDesc; col++; // 14
                worksheet.Cells[row, col].Value = cv.EwtAmt; col++; perCvEntryInputVatAmountTotal += cv.EwtAmt; // 15
                worksheet.Cells[row, col].Value = cv.EwtAcctNo; col++;
                worksheet.Cells[row, col].Value = cv.EwtDesc; col++;

                col = 27;
                worksheet.Cells[row, col].Value = cv.VatShouldBe; col++; vatShouldBeTaxTotal += cv.VatShouldBe; //27
                worksheet.Cells[row, col].Value = cv.VatAmt - cv.VatShouldBe; col++; vatVarianceonTaxTotal += (cv.VatAmt - cv.VatShouldBe); //28

                col = 34; // 29

                var cost = 0m;
                var expDr = 0m;
                var expCr = 0m;
                var capex = 0m;
                var vat = 0m;
                var def = 0m;
                var ewt = 0m;

                var selectedCvEntries = cventries
                    .Where(cv2 => cv2.CvNo!.Trim() == cv.Header.CvNo!.Trim())
                    .ToList();

                foreach(var selectedcventry in selectedCvEntries)
                {
                    if (selectedcventry.Acctcd.StartsWith("50"))
                    {
                        cost += selectedcventry.DrCr ? selectedcventry.Amount * -1 ?? 0m : selectedcventry.Amount ?? 0m;
                    }
                    if (selectedcventry.Acctcd.StartsWith("55") || selectedcventry.Acctcd.StartsWith("65"))
                    {
                        expDr += selectedcventry.DrCr ? 0m : selectedcventry.Amount ?? 0m;
                        expCr += selectedcventry.DrCr ? selectedcventry.Amount*-1 ?? 0m : 0m;
                    }
                    if (selectedcventry.Acctcd.StartsWith("102010"))
                    {
                        capex += selectedcventry.DrCr ? selectedcventry.Amount * -1 ?? 0m : selectedcventry.Amount ?? 0m;
                    }
                    if (selectedcventry.Acctcd.StartsWith("101060200"))
                    {
                        vat += selectedcventry.DrCr ? selectedcventry.Amount * -1 ?? 0m : selectedcventry.Amount ?? 0m;
                    }
                    if (selectedcventry.Acctcd.StartsWith("101060300"))
                    {
                        def += selectedcventry.DrCr ? selectedcventry.Amount * -1 ?? 0m : selectedcventry.Amount ?? 0m;
                    }
                    if (selectedcventry.Acctcd.StartsWith("201030"))
                    {
                        ewt += selectedcventry.DrCr ? selectedcventry.Amount * -1 ?? 0m : selectedcventry.Amount ?? 0m;
                    }
                }

                worksheet.Cells[row, col].Value = cost; col++; cost50Total += cost; //34
                worksheet.Cells[row, col].Value = expDr; col++; expenseDebitTotal += expDr; //35
                worksheet.Cells[row, col].Value = expCr; col++; expenseCreditTotal += expCr; //36
                worksheet.Cells[row, col].Value = capex; col++; expenseCapexTotal += capex; //37
                worksheet.Cells[row, col].Value = vat; col++; expenseVatTotal += vat; //38
                worksheet.Cells[row, col].Value = def; col++; expenseDVatTotal += def; //39
                worksheet.Cells[row, col].Value = ewt; col += 2; expenseEwtTotal += ewt; //40

                if(cv.DCRDate == null || cv.DCRDate > cv.DateTo)
                {
                    worksheet.Cells[row, col].Value = cv.Header.Amount; unclearedChecksTotal += cv.Header.Amount ?? 0m; //42
                }

                worksheet.Cells[row, 2].Style.Numberformat.Format = "dd-mmm-yyyy";
                worksheet.Cells[row, 6].Style.Numberformat.Format = "dd-mmm-yyyy";
                worksheet.Cells[row, 9].Style.Numberformat.Format = currencyFormat;
                worksheet.Cells[row, 12].Style.Numberformat.Format = currencyFormat;
                worksheet.Cells[row, 15].Style.Numberformat.Format = currencyFormat;
                worksheet.Cells[row, 27, row, 30].Style.Numberformat.Format = currencyFormat;
                worksheet.Cells[row, 32, row, 42].Style.Numberformat.Format = currencyFormat;

                row++;
            }

            #endregion == Values ==

            #region == Summary ==

            worksheet.Cells[row, 9].Value = vchAmountTotal;
            worksheet.Cells[row, 12].Value = perCvEntryAmountTotal;
            worksheet.Cells[row, 15].Value = perCvEntryInputVatAmountTotal;
            worksheet.Cells[row, 27].Value = vatShouldBeTaxTotal;
            worksheet.Cells[row, 28].Value = vatVarianceonTaxTotal;
            worksheet.Cells[row, 34].Value = cost50Total;
            worksheet.Cells[row, 35].Value = expenseDebitTotal;
            worksheet.Cells[row, 36].Value = expenseCreditTotal;
            worksheet.Cells[row, 37].Value = expenseCapexTotal;
            worksheet.Cells[row, 38].Value = expenseVatTotal;
            worksheet.Cells[row, 39].Value = expenseDVatTotal;
            worksheet.Cells[row, 40].Value = expenseEwtTotal;
            worksheet.Cells[row, 42].Value = unclearedChecksTotal;

            mergedCells = worksheet.Cells[row, 9, row, 42];
            mergedCells.Style.Numberformat.Format = currencyFormat;
            mergedCells.Style.Font.Bold = true;
            mergedCells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            mergedCells.Style.Border.Bottom.Style = ExcelBorderStyle.Double;

            #endregion == Summary

            #region == Cell sizes ==

            worksheet.View.FreezePanes(6, 1);
            worksheet.Columns.AutoFit();

            for (int ctr = 1; ctr != 45; ctr++)
            {
                if (worksheet.Column(ctr).Width < 15)
                {
                    worksheet.Column(ctr).Width = 15;
                }
                if (worksheet.Column(ctr).Width > 60)
                {
                    worksheet.Column(ctr).Width = 60;
                }
            }

            foreach (var colTemp in new List<int> { 11, 18, 26, 29, 33, 41 })
            {
                worksheet.Column(colTemp).Width = 1;
            }

            worksheet.Row(5).Height = 45;

            #endregion == Cell sizes ==
        }

        public void ProcessDetailReport(ExcelWorksheet worksheet, 
            GenerateTaxDeclarationViewModel viewModel, 
            List<CurcvheaderViewModel> curcvheader,
            List<CvEntriesViewModel> cventries,
            List<CvEntriesViewModel> cvEntries2)
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

            // var currencyFormat = "#,##0.00;[Red](#,##0.00)";
            var currencyFormat = "#,##0.00;[Red](#,##0.00);??;@";

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
            worksheet.Cells[row + 1, col].Value = "DR"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "CR"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "ACCT #"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "ACCT NAME"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col+=2;
            colSpandEnd = col - 2;

            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PeachPuff);

            
            colSpanStart = col;
            worksheet.Cells[row + 1, col].Value = "CATEGORY (OPEX/CAPEX)"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            colSpandEnd = col - 1;

            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

            worksheet.Cells[row, 1, row + 1, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, 1, row + 1, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[row, 1, row + 1, col].Style.Font.Bold = true;

            #endregion == Headers ==

            #region == Values ==

            var vchAmountTotal = 0m;
            var perCvEntryAmountTotal = 0m;
            var perCvEntryInputVatAmountTotal = 0m;
            var vatShouldBeTaxTotal = 0m;
            var vatVarianceonTaxTotal = 0m;
            var drTotal = 0m;
            var crTotal = 0m;

            row = 6;

            foreach (var cv in curcvheader)
            {
                col = 1;

                worksheet.Cells[row, col].Value = cv.Header!.Reference != null ? $"{cv.Header.CvNo} / {cv.Header.Reference}" : cv.Header.CvNo; col++;

                worksheet.Cells[row, col].Value = cv.Header.TranDate; col++; // 2
                worksheet.Cells[row, col].Value = cv.Header.Payee?.Trim(); col++; //3
                worksheet.Cells[row, col].Value = cv.Header.Particulars?.Trim(); col++; //4
                worksheet.Cells[row, col].Value = cv.Header.CheckNo?.Trim(); col++; //5
                worksheet.Cells[row, col].Value = cv.Header.CheckDate.ToString()?.Trim(); col++; // 6 date
                worksheet.Cells[row, col].Value = cv.Header.BankCode; col++; //7
                worksheet.Cells[row, col].Value = cv.DCRDate; col++; //8
                worksheet.Cells[row, col].Value = cv.Header.Amount; col++; vchAmountTotal += cv.Header.Amount ?? 0m; // 9
                worksheet.Cells[row, col].Value = cv.Header.BsNo; col++; // 10
                col++; // 11
                worksheet.Cells[row, col].Value = cv.VatAmt; col++; perCvEntryAmountTotal += cv.VatAmt; // 12
                worksheet.Cells[row, col].Value = cv.VatAcctNo; col++; // 13
                worksheet.Cells[row, col].Value = cv.VatDesc; col++; // 14
                worksheet.Cells[row, col].Value = cv.EwtAmt; col++; perCvEntryInputVatAmountTotal += cv.EwtAmt; // 15
                worksheet.Cells[row, col].Value = cv.EwtAcctNo; col++;
                worksheet.Cells[row, col].Value = cv.EwtDesc; col++;

                col = 27;
                worksheet.Cells[row, col].Value = cv.VatShouldBe; col++; vatShouldBeTaxTotal += cv.VatShouldBe; //27
                worksheet.Cells[row, col].Value = cv.VatAmt - cv.VatShouldBe; col++; vatVarianceonTaxTotal += (cv.VatAmt - cv.VatShouldBe); //28

                worksheet.Cells[row, 2].Style.Numberformat.Format = "dd-mmm-yyyy";
                worksheet.Cells[row, 6].Style.Numberformat.Format = "dd-mmm-yyyy";
                worksheet.Cells[row, 9].Style.Numberformat.Format = currencyFormat;
                worksheet.Cells[row, 12].Style.Numberformat.Format = currencyFormat;
                worksheet.Cells[row, 15].Style.Numberformat.Format = currencyFormat;
                worksheet.Cells[row, 27, row, 30].Style.Numberformat.Format = currencyFormat;
                worksheet.Cells[row, 32, row, 39].Style.Numberformat.Format = currencyFormat;

                var details = cvEntries2
                    .Where(cv2 => cv2.CvNo == cv.Header.CvNo)
                    .OrderBy(cv2 => cv2.Acctcd)
                    .ToList();

                foreach (var dtl in details)
                {
                    col = 34; // 

                    worksheet.Cells[row, col].Value = dtl.DrCr ? 0m : dtl.Amount; drTotal += dtl.DrCr ? 0m : dtl.Amount ?? 0m; col++; // 34
                    worksheet.Cells[row, col].Value = dtl.DrCr ? dtl.Amount*-1 : 0m ; crTotal += dtl.DrCr ? dtl.Amount * -1 ?? 0m : 0m; col++; // 35
                    worksheet.Cells[row, col].Value = dtl.Acctcd; col++; // 36
                    worksheet.Cells[row, col].Value = dtl.AcctName; col+=2; // 37

                    if (dtl.Acctcd.StartsWith("6"))
                    {
                        worksheet.Cells[row, col].Value = "OPEX"; // 38
                    }
                    if (dtl.Acctcd.StartsWith("105"))
                    {
                        worksheet.Cells[row, col].Value = "CAPEX"; // 38
                    }

                    worksheet.Cells[row, 34].Style.Numberformat.Format = currencyFormat;
                    worksheet.Cells[row, 35].Style.Numberformat.Format = currencyFormat;

                    row++;
                }

                row++;
            }

            #endregion == Values ==

            #region == Summary ==

            worksheet.Cells[row, 9].Value = vchAmountTotal;
            worksheet.Cells[row, 12].Value = perCvEntryAmountTotal;
            worksheet.Cells[row, 15].Value = perCvEntryInputVatAmountTotal;
            worksheet.Cells[row, 27].Value = vatShouldBeTaxTotal;
            worksheet.Cells[row, 28].Value = vatVarianceonTaxTotal;
            worksheet.Cells[row, 34].Value = drTotal;
            worksheet.Cells[row, 35].Value = crTotal;

            mergedCells = worksheet.Cells[row, 9, row, 35];
            mergedCells.Style.Numberformat.Format = currencyFormat;
            mergedCells.Style.Font.Bold = true;
            mergedCells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            mergedCells.Style.Border.Bottom.Style = ExcelBorderStyle.Double;

            #endregion == Summary

            #region == Cell sizes ==

            worksheet.View.FreezePanes(6, 1);
            worksheet.Columns.AutoFit();

            for (int ctr = 1; ctr != 39; ctr++)
            {
                if (worksheet.Column(ctr).Width < 15)
                {
                    worksheet.Column(ctr).Width = 15;
                }
                if (worksheet.Column(ctr).Width > 60)
                {
                    worksheet.Column(ctr).Width = 60;
                }
            }

            foreach (var colTemp in new List<int> { 11, 18, 26, 29, 33, 38 })
            {
                worksheet.Column(colTemp).Width = 1;
            }

            worksheet.Row(5).Height = 45;

            #endregion == Cell sizes ==
        }

        public void ProcessTrialBalanceReport(ExcelWorksheet worksheet, 
            GenerateTaxDeclarationViewModel viewModel,
            List<CurtrialbalViewModel> trialBal)
        {

            #region == Title Area ==

            var mergedCells = worksheet.Cells["A1:C1"];
            mergedCells.Merge = true;
            mergedCells.Value = $"IBS DISBURSEMENT VOUCHERS - {viewModel.SelectedCompany} - TRIAL BALANCE";
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

            var currencyFormat = "#,##0.00;[Red](#,##0.00);??;@";

            #region == Headers ==

            var row = 4;
            var col = 1;

            var colSpanStart = col;
            worksheet.Cells[row, col].Value = "ACCOUNT#"; worksheet.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row, col].Value = "ACCOUNT NAME"; worksheet.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row, col].Value = "DEBIT"; worksheet.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row, col].Value = "CREDIT"; worksheet.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            var colSpandEnd = col - 1;

            worksheet.Cells[row, colSpanStart, row, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, colSpanStart, row, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PeachPuff);

            worksheet.Cells[row, 1, row, 42].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, 1, row, 42].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[row, 1, row, 42].Style.Font.Bold = true;

            #endregion == Headers ==

            #region == Values ==

            var debitTotal = 0m;
            var creditTotal = 0m;

            row = 5;

            foreach (var tb in trialBal)
            {
                col = 1;

                worksheet.Cells[row, col].Value = tb.Acctcd?.Trim(); col++;
                worksheet.Cells[row, col].Value = tb.AcctName?.Trim(); col++;
                worksheet.Cells[row, col].Value = tb.Bal > 0 ? tb.Bal : null; col++;
                worksheet.Cells[row, col].Value = tb.Bal < 0 ? tb.Bal : null; col++;

                if (tb.Bal > 0)
                {
                    debitTotal += tb.Bal ?? 0m;
                }
                if (tb.Bal < 0)
                {
                    creditTotal += tb.Bal ?? 0m;
                }

                worksheet.Cells[row, 3, row, 4].Style.Numberformat.Format = currencyFormat;

                row++;
            }

            #endregion == Values ==

            #region == Summary ==

            worksheet.Cells[row, 3].Value = debitTotal;
            worksheet.Cells[row, 4].Value = creditTotal;

            worksheet.Cells[row, 3, row, 4].Style.Numberformat.Format = currencyFormat;
            worksheet.Cells[row, 3, row, 4].Style.Font.Bold = true;
            worksheet.Cells[row, 3, row, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[row, 3, row, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Double;

            #endregion == Summary

            #region == Cell sizes ==

            worksheet.View.FreezePanes(5, 1);
            worksheet.Columns.AutoFit();

            for (int ctr = 1; ctr != 4; ctr++)
            {
                if (worksheet.Column(ctr).Width < 15)
                {
                    worksheet.Column(ctr).Width = 15;
                }
                if (worksheet.Column(ctr).Width > 60)
                {
                    worksheet.Column(ctr).Width = 60;
                }
            }


            #endregion == Cell sizes ==
        }

        public void ProcessGeneralLedgerReport(ExcelWorksheet worksheet,
            GenerateTaxDeclarationViewModel viewModel,
            List<CvEntriesViewModel> cventries2,
            bool isGLAll)
        {

            #region == Title Area ==

            var mergedCells = worksheet.Cells["A1:C1"];
            mergedCells.Merge = true;
            mergedCells.Value = $"IBS DISBURSEMENT VOUCHERS - {viewModel.SelectedCompany} - GENERAL LEDGER";
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

            var currencyFormat = "#,##0.00;[Red](#,##0.00);??;@";

            #region == Headers ==

            var row = 4;
            var col = 1;

            var colSpanStart = col;
            worksheet.Cells[row + 1, col].Value = "VOUCHER#"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "VOUCHER DATE"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "PAYEE"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "PARTICULAR"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "CHECK#"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "CHECK DATE"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "OTHERS"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "VCH AMOUNT"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "DEBIT"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "CREDIT"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "ACCOUNT#"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
            worksheet.Cells[row + 1, col].Value = "ACCOUNT NAME"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col += 2;
            var colSpandEnd = col - 2;

            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PeachPuff);

            if (isGLAll)
            {
                colSpanStart = col;
                mergedCells = worksheet.Cells[row, col, row, col + 6]; mergedCells.Merge = true; mergedCells.Value = "FROM 2307 SYSTEM"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                mergedCells.Style.Fill.PatternType = ExcelFillStyle.Solid; mergedCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                worksheet.Cells[row + 1, col].Value = "PAYOR"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
                worksheet.Cells[row + 1, col].Value = "PAYEE"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
                worksheet.Cells[row + 1, col].Value = "TIN"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
                worksheet.Cells[row + 1, col].Value = "PERCENT"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
                worksheet.Cells[row + 1, col].Value = "EWT AMOUNT"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
                worksheet.Cells[row + 1, col].Value = "MONTH"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col++;
                worksheet.Cells[row + 1, col].Value = "YEAR"; worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col = col + 2;
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
                mergedCells = worksheet.Cells[row, col, row + 1, col]; mergedCells.Merge = true; mergedCells.Value = "EWT"; mergedCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black); col = col + 2;
                colSpandEnd = col - 2;

                worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, colSpanStart, row + 1, colSpandEnd].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                worksheet.Cells[row, 1, row + 1, 28].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[row, 1, row + 1, 28].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[row, 1, row + 1, 28].Style.Font.Bold = true;
            }
            else
            {
                worksheet.Cells[row, 1, row + 1, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[row, 1, row + 1, 12].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[row, 1, row + 1, 12].Style.Font.Bold = true;
            }

            #endregion == Headers ==

            #region == Values ==

            var debitTotal = 0m;
            var creditTotal = 0m;

            var cost50Total = 0m;
            var expenseDebitTotal = 0m;
            var expenseCreditTotal = 0m;
            var expenseCapexTotal = 0m;
            var expenseVatTotal = 0m;
            var expenseDVatTotal = 0m;
            var expenseEwtTotal = 0m;
            var unclearedChecksTotal = 0m;

            row = 6;

            foreach (var cv in cventries2)
            {
                col = 1;

                worksheet.Cells[row, col].Value = cv.CvNo; col++;
                worksheet.Cells[row, col].Value = cv.TranDate; col++;
                worksheet.Cells[row, col].Value = cv.Payee; col++;
                worksheet.Cells[row, col].Value = cv.Particulars; col++;
                worksheet.Cells[row, col].Value = cv.CheckNo; col++;
                worksheet.Cells[row, col].Value = cv.CheckDate; col++;
                worksheet.Cells[row, col].Value = cv.BsNo; col++;
                worksheet.Cells[row, col].Value = cv.CvAmount; col++;
                worksheet.Cells[row, col].Value = cv.DrCr ? "" : cv.Amount ?? 0m; col++;
                worksheet.Cells[row, col].Value = cv.DrCr ? cv.Amount * -1 : ""; col++;
                worksheet.Cells[row, col].Value = cv.Acctcd; col++;
                worksheet.Cells[row, col].Value = cv.AcctName; col++;

                if (isGLAll)
                {
                    col = 22;

                    var cost = 0m;
                    var expDr = 0m;
                    var expCr = 0m;
                    var capex = 0m;
                    var vat = 0m;
                    var def = 0m;
                    var ewt = 0m;

                    if (cv.Acctcd.StartsWith("50"))
                    {
                        cost = cv.DrCr ? cv.Amount * -1 ?? 0m : cv.Amount ?? 0m;
                        cost50Total += cost;
                    }
                    if (cv.Acctcd.StartsWith("55") || cv.Acctcd.StartsWith("65"))
                    {
                        expDr = cv.DrCr ? 0m : cv.Amount ?? 0m;
                        expCr = cv.DrCr ? cv.Amount * -1 ?? 0m : 0m;

                        expenseDebitTotal += expDr;
                        expenseCreditTotal += expCr;
                    }
                    if (cv.Acctcd.StartsWith("102010"))
                    {
                        capex += cv.DrCr ? cv.Amount * -1 ?? 0m : cv.Amount ?? 0m;
                        expenseCapexTotal += capex;
                    }
                    if (cv.Acctcd.StartsWith("101060200"))
                    {
                        vat += cv.DrCr ? cv.Amount * -1 ?? 0m : cv.Amount ?? 0m;
                        expenseVatTotal += vat;
                    }
                    if (cv.Acctcd.StartsWith("101060300"))
                    {
                        def += cv.DrCr ? cv.Amount * -1 ?? 0m : cv.Amount ?? 0m;
                        expenseDVatTotal += def;
                    }
                    if (cv.Acctcd.StartsWith("201030"))
                    {
                        ewt += cv.DrCr ? cv.Amount * -1 ?? 0m : cv.Amount ?? 0m;
                        expenseEwtTotal += ewt;
                    }

                    worksheet.Cells[row, col].Value = cost; col++;
                    worksheet.Cells[row, col].Value = expDr; col++;
                    worksheet.Cells[row, col].Value = expCr; col++;
                    worksheet.Cells[row, col].Value = capex; col++;
                    worksheet.Cells[row, col].Value = vat; col++;
                    worksheet.Cells[row, col].Value = def; col++;
                    worksheet.Cells[row, col].Value = ewt; col += 2;

                    worksheet.Cells[row, 22, row, 28].Style.Numberformat.Format = currencyFormat;
                }

                if (cv.DrCr)
                {
                    creditTotal += (cv.Amount ?? 0m) * -1;
                }
                else
                {
                    debitTotal += cv.Amount ?? 0m;
                }

                worksheet.Cells[row, 9, row, 10].Style.Numberformat.Format = currencyFormat;

                row++;
            }

            #endregion == Values ==

            #region == Summary ==

            worksheet.Cells[row, 9].Value = debitTotal;
            worksheet.Cells[row, 10].Value = creditTotal;

            worksheet.Cells[row, 9, row, 10].Style.Numberformat.Format = currencyFormat;
            worksheet.Cells[row, 9, row, 10].Style.Font.Bold = true;
            worksheet.Cells[row, 9, row, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[row, 9, row, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Double;

            #endregion == Summary ==

            #region == Summary ==

            worksheet.Cells[row, 22].Value = cost50Total;
            worksheet.Cells[row, 23].Value = expenseDebitTotal;
            worksheet.Cells[row, 24].Value = expenseCreditTotal;
            worksheet.Cells[row, 25].Value = expenseCapexTotal;
            worksheet.Cells[row, 26].Value = expenseVatTotal;
            worksheet.Cells[row, 27].Value = expenseDVatTotal;
            worksheet.Cells[row, 28].Value = expenseEwtTotal;
            worksheet.Cells[row, 29].Value = unclearedChecksTotal;

            mergedCells = worksheet.Cells[row, 9, row, 42];
            mergedCells.Style.Numberformat.Format = currencyFormat;
            mergedCells.Style.Font.Bold = true;
            mergedCells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            mergedCells.Style.Border.Bottom.Style = ExcelBorderStyle.Double;

            #endregion == Summary

            #region == Cell sizes ==

            worksheet.View.FreezePanes(6, 1);
            worksheet.Columns.AutoFit();

            for (int ctr = 1; ctr != 28; ctr++)
            {
                if(worksheet.Column(ctr).Width < 15)
                {
                    worksheet.Column(ctr).Width = 15;
                }
                if(worksheet.Column(ctr).Width > 60)
                {
                    worksheet.Column(ctr).Width = 60;
                }
            }

            #endregion == Cell sizes ==
        }
    }
}
