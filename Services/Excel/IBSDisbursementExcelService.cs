using OfficeOpenXml;
using OfficeOpenXml.Style;
using TaxDeclaration.Models.ViewModels;

namespace TaxDeclaration.Services.Excel
{
    public class IBSDisbursementExcelService
    {
        public void ProcessReport1(ExcelWorksheet worksheet, 
            GenerateTaxDeclarationViewModel viewModel, 
            List<CurcvheaderViewModel> curcvheader,
            List<CvEntriesViewModel> cventries)
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

            var currencyFormat = "#,##0.00";

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

            #region == Values ==

            row = 6;

            foreach (var cv in curcvheader)
            {
                col = 1;

                worksheet.Cells[row, col].Value = cv.Header!.Reference != null ? $"{cv.Header.CvNo} / {cv.Header.Reference}" : cv.Header.CvNo; col++;

                worksheet.Cells[row, col].Value = cv.Header.TranDate; col++; // 2 date
                worksheet.Cells[row, col].Value = cv.Header.Payee?.Trim(); col++; //3
                worksheet.Cells[row, col].Value = cv.Header.Particulars?.Trim(); col++; //4
                worksheet.Cells[row, col].Value = cv.Header.CheckNo?.Trim(); col++;//5
                worksheet.Cells[row, col].Value = cv.Header.CheckDate.ToString()?.Trim(); col++; // 6 date
                worksheet.Cells[row, col].Value = cv.Header.BankCode; col++;//7

                //worksheet.Cells[row, col].Value = cv.AccountNo; col++;
                //worksheet.Cells[row, col].Value = cv.CashpoDate; col++;

                worksheet.Cells[row, col].Value = cv.DCRDate; col++;//8
                worksheet.Cells[row, col].Value = cv.Header.Amount; col++; // 9
                worksheet.Cells[row, col].Value = cv.Header.BsNo; col++; // 10
                col++; // 11
                worksheet.Cells[row, col].Value = cv.VatAmt; col++; // 12
                worksheet.Cells[row, col].Value = cv.VatAcctNo; col++;
                worksheet.Cells[row, col].Value = cv.VatDesc; col++;
                worksheet.Cells[row, col].Value = cv.EwtAmt; col++;
                worksheet.Cells[row, col].Value = cv.EwtAcctNo; col++;
                worksheet.Cells[row, col].Value = cv.EwtDesc; col++;

                col = 27;
                worksheet.Cells[row, col].Value = cv.VatShouldBe; col++; //27
                worksheet.Cells[row, col].Value = cv.VatAmt - cv.VatShouldBe; col++; //28

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

                worksheet.Cells[row, col].Value = cost; col++; //34
                worksheet.Cells[row, col].Value = expDr; col++; //35
                worksheet.Cells[row, col].Value = expCr; col++; //36
                worksheet.Cells[row, col].Value = capex; col++; //37
                worksheet.Cells[row, col].Value = vat; col++; //38
                worksheet.Cells[row, col].Value = def; col++; //39
                worksheet.Cells[row, col].Value = ewt; col += 2; //40

                if(cv.DCRDate == null || cv.DCRDate > cv.DateTo)
                {
                    worksheet.Cells[row, col].Value = cv.Header.Amount; //42
                }

                worksheet.Cells[row, 2].Style.Numberformat.Format = "dd-mmm-yyyy";
                worksheet.Cells[row, 6].Style.Numberformat.Format = "dd-mmm-yyyy";
                worksheet.Cells[row, 9].Style.Numberformat.Format = currencyFormat;
                worksheet.Cells[row, 34, row, 42].Style.Numberformat.Format = currencyFormat;

                row++;
            }

            #endregion == Values ==

            #region == Cell sizes ==

            worksheet.View.FreezePanes(6, 1);
            worksheet.Columns.AutoFit();

            for (int ctr = 1; ctr != 45; ctr++)
            {
                worksheet.Column(ctr).Width = 25;
            }

            foreach (var colTemp in new List<int> { 11, 18, 26, 29, 33, 41 })
            {
                worksheet.Column(colTemp).Width = 1;
            }

            worksheet.Row(5).Height = 45;

            #endregion == Cell sizes ==
        }
    }
}
