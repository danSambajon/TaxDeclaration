using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Reflection.PortableExecutable;
using System.Runtime.ConstrainedExecution;
using TaxDeclaration.Models;
using TaxDeclaration.Models.ViewModels;
using TaxDeclaration.Services;
using TaxDeclaration.Services.Dbf;
using TaxDeclaration.Services.Excel;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;

namespace TaxDeclaration.Controllers
{
    public class IBSController : Controller
    {
        private readonly DbfService _dbfService;
        private readonly IBSConnectionService _ibsConnectionService;
        private readonly IBSDisbursementExcelService _excel;

        public IBSController(DbfService dbfService, IBSConnectionService ibsConnectionService, IBSDisbursementExcelService excel)
        {
            _dbfService = dbfService;
            _ibsConnectionService = ibsConnectionService;
            _excel = excel;
        }

        public IActionResult Index()
        {
            var viewModel = new GenerateTaxDeclarationViewModel
            {
                Company = "FILPRIDE",
                CompanyChoices = new List<SelectListItem>
                {
                    new SelectListItem { Text = "FILPRIDE", Value = "FILPRIDE" }
                }
            };

            viewModel.SelectedCompany = viewModel.CompanyChoices.FirstOrDefault()!.Value;
            return View(viewModel);
        }

        public async Task<IActionResult> Process(GenerateTaxDeclarationViewModel viewModel, CancellationToken cancellationToken)
        {
            // Default file for IBS tax declaration
            viewModel.FileFormat = "Excel";

            #region == Initialize data containers ==

            var stations = new List<FastStationViewModel>();
            var fastCvs = new List<FastCvViewModel>();
            var companies = new List<TaxDeclareCompanyViewModel>();
            var dcrMainDisbursements = new List<DCRMainDisbursementViewModel>();
            var cur2307 = new List<Cur2307ViewModel>();

            var checkVoucherHeaders = new List<FilprideCheckVoucherHeader>();
            var checkVoucherDetails = new List<FilprideCheckVoucherDetail>();
            var chartOfAccounts = new List<FilprideChartOfAccount>();

            #endregion == Initialize data containers ==

            try
            {
                #region == Data related ==

                //#region == Getting data from sources ==

                //// Get data from DBF files
                //stations = _dbfService.GetStationsFromDbf();
                //fastCvs = _dbfService.GetFastCvsFromDbf();
                //companies = _dbfService.GetCompaniesFromDbf();
                //dcrMainDisbursements = _dbfService.GetDCRMainDisbursementsFromDbf();
                //cur2307 = _dbfService.GetBirSysFromDbf();

                //// Get data from IBS
                //checkVoucherHeaders = await _ibsConnectionService
                //    .GetCheckVoucherHeaders(viewModel.DateFrom, viewModel.DateTo);

                //if (checkVoucherHeaders.Count == 0)
                //{
                //    throw new NullReferenceException("No check voucher headers found for the specified date range.");
                //}

                //checkVoucherDetails = await _ibsConnectionService
                //    .GetCheckVoucherDetails(checkVoucherHeaders.Select(h => h.CheckVoucherHeaderNo).ToList(), checkVoucherHeaders.Select(h => h.Reference).ToList());

                //chartOfAccounts = await _ibsConnectionService
                //    .GetChartOfAccounts();

                //#endregion == Getting data from sources ==

                //#region == Cursors ==

                //foreach (FilprideCheckVoucherHeader header in checkVoucherHeaders)
                //{
                //    header.Details = checkVoucherDetails.Where(d => d.TransactionNo == header.CheckVoucherHeaderNo).ToList();
                //    header.Referenced = checkVoucherDetails.Where(d => d.TransactionNo == header.Reference).ToList();
                //}

                //foreach (FilprideCheckVoucherDetail detail in checkVoucherDetails)
                //{
                //    detail.Header = checkVoucherHeaders.Where(h => h.CheckVoucherHeaderNo == detail.TransactionNo).FirstOrDefault();

                //    if (detail.Header != null)
                //    {
                //        detail.CompanyVm = companies.Where(c => c.BankCode == detail.Header.BankAccountNumber).FirstOrDefault();
                //    }
                //}

                //// Verified
                //var cvHeaderx = checkVoucherHeaders
                //    .Select(h => new CvheaderxViewModel
                //    {
                //        CvNo = h.CheckVoucherHeaderNo,
                //        TranDate = h.Date,
                //        Payee = h.Payee,
                //        BankCode = h.BankAccountNumber,
                //        BankName = h.BankAccountNumber,
                //        CheckNo = h.CheckNo,
                //        Amount = h.Total,
                //        CheckDate = h.CheckDate,
                //        Particulars = h.Particulars,
                //        CreatedBy = h.CreatedBy,
                //        AcctCd = null,
                //        BsNo = h.Type?.ToUpper(),
                //        CheckClearing = h.DcrDate,
                //        NameCategory = null,
                //        IsCancelled = h.CanceledBy != null,
                //        ChkClear = h.DcrDate,
                //        Type = h.Type,
                //        Reference = h.Reference,
                //        Category = h.Category,
                //        CvType = h.CvType
                //    })
                //    .ToList();

                //var cvNoFromcvheaderx = cvHeaderx.Select(h => h.CvNo).ToList();
                //var referenceFromcvheaderx = cvHeaderx.Select(h => h.Reference).ToList();

                //// Verified
                //var cvDetailx = checkVoucherDetails
                //    .Where(d => !d.IsDisplayEntry && (cvNoFromcvheaderx.Contains(d.TransactionNo!.Trim()) || referenceFromcvheaderx.Contains(d.TransactionNo!.Trim())))
                //    .Select(d => new CvdetailxViewModel
                //    {
                //        CvNo = d.TransactionNo,
                //        SeqId = d.CheckVoucherDetailId,
                //        Amount = d.Debit.GetValueOrDefault() != 0 ? d.Debit : d.Credit,
                //        DrCr = d.Debit.GetValueOrDefault() == 0,
                //        Acctcd = d.AccountNo,
                //        IsDisplayEntry = d.IsDisplayEntry,
                //        Header = d.Header
                //    })
                //    .ToList();

                //var bankCodesOfCvDetailx = cvDetailx.Select(d => d.Header.BankAccountNumber).Distinct().ToList();

                //// Verified
                //// DETAILS: VAT INPUT AND GOVERNMENT PAYABLES 
                //var temp = cvDetailx
                //    .Where(d => d.Acctcd.StartsWith("101060200") || d.Acctcd.StartsWith("201030"))
                //    .OrderBy(d => d.CvNo)
                //    .Select(d => d.CvNo)
                //    .Distinct()
                //    .ToList();

                //// Verified
                //// DETAILS(GROUPED): VAT INPUT AND GOVERNMENT PAYABLES
                //var temp2 = cvDetailx
                //    .Where(d => d.Acctcd.StartsWith("101060200") || d.Acctcd.StartsWith("201030"))
                //    .GroupBy(d => d.CvNo)
                //    .Select(g => new
                //    {
                //        CvNo = g.Key,
                //        Ctr = g.Count(),
                //        Records = g.ToList(),
                //        Header = cvHeaderx.Where(h => h.CvNo == g.Key).FirstOrDefault(),
                //        Company = companies.Where(c => c.BankCode == cvHeaderx.Where(h => h.CvNo == g.Key).FirstOrDefault().BankCode).FirstOrDefault()
                //    })
                //    .OrderByDescending(g => g.Ctr)
                //    .ToList();

                //// Verified
                //// DETAILS + HEADER: WITH EMPTY FIELDS
                //var curcvheader = temp2
                //    .Select(d => new CurcvheaderViewModel
                //    {
                //        CvDtl = d.CvNo,
                //        Header = d.Header,
                //        CashpoDate = new DateOnly(),
                //        DCRDate = new DateOnly(),
                //        Company = d.Company == null ? (string)null : d.Company.Co
                //    })
                //    .Where(d => d.Header.TranDate >= viewModel.DateFrom && d.Header.TranDate <= viewModel.DateTo)
                //    .ToList();

                //// Verified
                //var cvDetailGroup = cvDetailx
                //    .Where(d => temp.Contains(d.CvNo))
                //    .GroupBy(d => new
                //    {
                //        d.CvNo,
                //        d.DrCr,
                //        d.Acctcd
                //    })
                //    .Select(g => new
                //    {
                //        g.Key.CvNo,
                //        g.Key.DrCr,
                //        g.Key.Acctcd,
                //        Amount = g.Sum(x => x.Amount),
                //        NEntry = g.Count(),
                //        SeqId = g.Max(x => x.SeqId)
                //    })
                //    .OrderBy(x => x.CvNo)
                //    .ThenBy(x => x.SeqId)
                //    .ToList();

                //// Verified
                //var cvDetailGroup2 = cvDetailx
                //    .Where(d => temp.Contains(d.CvNo))
                //    .OrderBy(d =>  d.CvNo )
                //    .ThenBy(d => d.SeqId)
                //    .Select(d => new
                //    {
                //        d.CvNo,
                //        d.DrCr,
                //        d.Acctcd,
                //        AcctName = chartOfAccounts.Where(coa => coa.AccountNumber == d.Acctcd).FirstOrDefault().AccountName,
                //        d.Amount,
                //        d.SeqId
                //    })
                //    .ToList();

                //// Verified
                //var detail_accounts = cvDetailGroup
                //    .Select(d => new DetailAccountsViewModel
                //    {
                //        DrCr = d.DrCr,
                //        Acctcd = d.Acctcd,
                //        AcctName = chartOfAccounts.Where(coa => coa.AccountNumber == d.Acctcd).FirstOrDefault().AccountName,
                //        ColNum = (decimal)0
                //    })
                //    .Distinct()
                //    .OrderBy(d => d.Acctcd)
                //    .ToList();

                //// Verified
                //var acctgroup = detail_accounts
                //    .GroupBy(d => new
                //    {
                //        d.DrCr,
                //        d.Acctcd
                //    })
                //    .Select(d => new
                //    {
                //        d.Key.DrCr,
                //        Acc = d.Key.Acctcd,
                //        Ctr = d.Count(),
                //        Colnum = (decimal)0
                //    })
                //    .ToList();

                //// Verified
                //var cventries = (from d in curcvheader
                //    join h in cvDetailGroup on d.Header.CvNo equals h.CvNo
                //    join a in detail_accounts
                //        on new { AcctCd = h.Acctcd.Trim(), DrCr = h.DrCr }
                //        equals new { AcctCd = a.Acctcd.Trim(), DrCr = a.DrCr }
                //        into leftJoin
                //    from c in leftJoin.DefaultIfEmpty()
                //    select new CvEntriesViewModel
                //    {
                //        CvNo = d.Header.CvNo,
                //        TranDate = d.Header.TranDate,
                //        Payee = d.Header.Payee,
                //        BankCode = d.Header.BankCode,
                //        BankName = d.Header.BankName,
                //        CvAmount = d.Header.Amount,
                //        CheckNo = d.Header.CheckNo,
                //        CheckDate = d.Header.CheckDate,
                //        BsNo = d.Header.BsNo,
                //        ChkClear = d.Header.ChkClear,
                //        Particulars = d.Header.Particulars,
                //        Amount = h.Amount,
                //        DrCr = h.DrCr,
                //        Acctcd = h.Acctcd,
                //        AcctName = c.AcctName,
                //        ColNum = c.ColNum,
                //        Reference = d.Header.Reference
                //    })
                //    .OrderBy(h => h.CvNo)
                //    .ThenBy(h => h.AcctName)
                //    .ToList();

                //// Verified
                //var cvEntries2 = (from a in curcvheader
                //    from b in cvDetailGroup2
                //    where a.Header.CvNo == b.CvNo || a.Header.Reference == b.CvNo
                //    join c in detail_accounts
                //        on new { AcctCd = b.Acctcd.Trim(), DrCr = b.DrCr }
                //        equals new { AcctCd = c.Acctcd.Trim(), DrCr = c.DrCr }
                //        into leftJoin
                //    from c in leftJoin.DefaultIfEmpty()
                //    select new CvEntriesViewModel
                //    {
                //        CvNo = a.Header.CvNo,
                //        TranDate = a.Header.TranDate,
                //        Payee = a.Header.Payee,
                //        BankCode = a.Header.BankCode,
                //        BankName = a.Header.BankName,
                //        CvAmount = a.Header.Amount,
                //        CheckNo = a.Header.CheckNo,
                //        CheckDate = a.Header.CheckDate,
                //        BsNo = a.Header.BsNo,
                //        ChkClear = a.Header.ChkClear,
                //        Particulars = a.Header.Particulars,
                //        Reference = a.Header.Reference,
                //        Category = a.Header.Category,
                //        Amount = b.Amount,
                //        DrCr = b.DrCr,
                //        Acctcd = b.Acctcd,
                //        Seqid = b.SeqId,
                //        AcctName = c != null ? c.AcctName : null,
                //        ColNum = c != null ? c.ColNum : 0,
                //    })
                //    .OrderBy(x => x.CvNo)
                //    .ThenBy(x => x.ColNum)
                //    .ToList();

                //#endregion == Cursors ==

                //#region == Assign DCR and 2307 values to cv ==

                //// DCR
                //foreach (var cv in curcvheader)
                //{
                //    var dcrEntry = dcrMainDisbursements.Where(dcr => dcr.VoucherNo.Trim() == cv.Header.CvNo.Trim()).FirstOrDefault();

                //    if (dcrEntry != null)
                //    {
                //        cv.CashpoDate = dcrEntry.CashPoDate;
                //        cv.DCRDate = dcrEntry.DcrDate;
                //        cv.AccountNo = dcrEntry.AccountNo;
                //        cv.StnCode = dcrEntry.StnCode;
                //        if(dcrEntry.AccountNo != cv.Header.BankCode)
                //        {
                //            cv.Rem = "DIFFERENT BANK CODE";
                //        }
                //    }
                //    else
                //    {
                //        cv.CashpoDate = null;
                //        cv.DCRDate = null;
                //    }
                //}

                //// 2307
                //foreach(var cv in curcvheader)
                //{
                //    var cur = cur2307.Where(c => c.CvNo.Trim() == cv.Header.CvNo.Trim() && c.DownloadFrom.Contains("IBS")).FirstOrDefault();

                //    if (cur != null)
                //    {
                //        cv.DateFrom = cur.DateFrom;
                //        cv.DateTo = cur.DateTo;
                //        cv.AtcCode = cur.AtcCode;
                //        cv.AtcDesc = cur.Desc;
                //        cv.PayeeName = cur.PayeeName;
                //        cv.PayorName = cur.PayorName;
                //        cv.Tin = $"{cur.TinA}-{cur.TinB}-{cur.TinC}-{cur.TinD}";
                //        cv.Percent = cur.Percent;
                //        cv.Total = cur.Total;
                //        cv.Amt2307 = cur.Amount;
                //        cv.NMonth = cur.RMonth;
                //        cv.NYear = cur.RYear;
                //    }
                //}

                //#endregion == Assign DCR and 2307 values to cv ==

                //#region == EWT and VAT ==

                //foreach (var cv in curcvheader)
                //{
                //    var vat = 0m;

                //    if((cv.Amt2307.HasValue && cv.Amt2307 != 0) && (cv.Percent.HasValue && cv.Percent!= 0))
                //    {
                //        cv.EwtShouldBe = cv.Amt2307/(cv.Percent/100);
                //    }

                //    var cvEntry = cventries
                //        .Where(c => c.CvNo == cv.Header.CvNo 
                //        && c.Acctcd.Trim() == "101060200" 
                //        && c.Acctcd.Trim() == "V00-17-101")
                //        .FirstOrDefault();

                //    // VAT
                //    if (cvEntry != null)
                //    {
                //        vat = (cvEntry.Amount ?? 0m) / 0.12m;

                //        cv.EwtAmt = cvEntry.Amount;
                //        cv.EwtAcctNo = cvEntry.Acctcd;
                //        cv.EwtDesc = cvEntry.AcctName;
                //        cv.EwtOk = true;
                //        cv.VatShouldBe = vat;
                //    }
                //    else
                //    {
                //        if (cv.EwtShouldBe.HasValue && cv.EwtShouldBe != 0)
                //        {
                //            var ewt = cv.EwtShouldBe;
                            
                //            var cvEntry2 = cventries
                //                .Where(c => c.CvNo.Trim() == cv.Header.CvNo.Trim() && c.Amount >= ewt-1 && c.Amount <= ewt+1)
                //                .FirstOrDefault();

                //            if(cvEntry2 != null)
                //            {
                //                cv.VatAmt = cvEntry2.Amount;
                //                cv.VatAcctNo = cvEntry2.Acctcd;
                //                cv.VatDesc = cvEntry2.AcctName;
                //                cv.VatOk = true;
                //            }
                //        }
                //    }

                //    // credit included ewt
                //    if (!cv.VatAmt.HasValue && cv.VatAmt == 0)
                //    {
                //        var amt = 0m;
                //        var ctr = 0m;
                //        var acct = string.Empty;
                //        var name = string.Empty;

                //        var acctList = new[] {
                //            "101060200",
                //            "101060300",
                //            "101010100"
                //        };

                //        var cvEntriesTemp = cvEntries2
                //            .Where(c => c.CvNo == cv.Header.CvNo
                //            && (acctList.Contains(c.Acctcd.Trim()) || c.Acctcd.Trim().StartsWith("201030"))
                //            && !c.Acctcd.Trim().StartsWith("V00")).ToList();

                //        if (cvEntriesTemp != null)
                //        {
                //            foreach (var cvEntryTemp in cvEntriesTemp)
                //            {
                //                amt = amt + (cvEntryTemp.DrCr ? (cvEntryTemp.Amount * -1) : cvEntryTemp.Amount) ?? 0m;
                //                ctr = ctr + 1;

                //                if (acct == string.Empty)
                //                {
                //                    acct = cvEntryTemp.Acctcd;
                //                    name = cvEntryTemp.AcctName;
                //                }

                //                if (acct != cvEntryTemp.Acctcd)
                //                {
                //                    acct = "SUM OF ";
                //                }
                //            }
                //        }

                //        if (cv.EwtShouldBe >= amt - 5 && cv.EwtShouldBe <= amt + 5)
                //        {
                //            if (acct == "SUM OF ")
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = "MULTIPLE";
                //                cv.VatDesc = $"SUM OF {ctr.ToString()} ENTRIES";
                //                cv.VatOk = true;
                //            }
                //            else
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = acct;
                //                cv.VatDesc = name;
                //                cv.VatOk = true;
                //            }
                //        }
                //    }
                     
                //    // credit excluded ewt
                //    if (!cv.VatAmt.HasValue && cv.VatAmt == 0)
                //    {
                //        var amt = 0m;
                //        var ctr = 0m;
                //        var acct = string.Empty;
                //        var name = string.Empty;

                //        var acctList = new[] {
                //            "101060200",
                //            "101060300",
                //            "101010100"
                //        };

                //        var cvEntriesTemp = cvEntries2
                //            .Where(c => c.CvNo == cv.Header.CvNo
                //            && (acctList.Contains(c.Acctcd.Trim()) || c.Acctcd.Trim().StartsWith("201030"))
                //            && !c.Acctcd.Trim().StartsWith("V00")).ToList();

                //        if (cvEntriesTemp != null)
                //        {
                //            foreach (var cvEntryTemp in cvEntriesTemp)
                //            {
                //                amt = amt + (cvEntryTemp.DrCr ? 0 : cvEntryTemp.Amount) ?? 0m;
                //                ctr = ctr + 1;

                //                if (acct == string.Empty)
                //                {
                //                    acct = cvEntryTemp.Acctcd;
                //                    name = cvEntryTemp.AcctName;
                //                }

                //                if (acct != cvEntryTemp.Acctcd)
                //                {
                //                    acct = "SUM OF ";
                //                }
                //            }
                //        }

                //        if (cv.EwtShouldBe >= amt - 5 && cv.EwtShouldBe <= amt + 5)
                //        {
                //            if (acct == "SUM OF ")
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = "MULTIPLE";
                //                cv.VatDesc = $"SUM OF {ctr.ToString()} ENTRIES";
                //                cv.VatOk = true;
                //            }
                //            else
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = acct;
                //                cv.VatDesc = name;
                //                cv.VatOk = true;
                //            }
                //        }
                //    }

                //    // credit included (2) vat
                //    if (!cv.VatAmt.HasValue && cv.VatAmt == 0)
                //    {
                //        var amt = 0m;
                //        var ctr = 0m;
                //        var acct = string.Empty;
                //        var name = string.Empty;

                //        var acctList = new[] {
                //            "101060200",
                //            "101060300",
                //            "101010100"
                //        };

                //        var cvEntriesTemp = cvEntries2
                //            .Where(c => c.CvNo == cv.Header.CvNo
                //            && (acctList.Contains(c.Acctcd.Trim()) || c.Acctcd.Trim().StartsWith("201030"))
                //            && !c.Acctcd.Trim().StartsWith("V00")).ToList();

                //        if (cvEntriesTemp != null)
                //        {
                //            foreach (var cvEntryTemp in cvEntriesTemp)
                //            {
                //                amt = amt + (cvEntryTemp.DrCr ? cvEntryTemp.Amount*-1 : cvEntryTemp.Amount) ?? 0m;
                //                ctr = ctr + 1;

                //                if (acct == string.Empty)
                //                {
                //                    acct = cvEntryTemp.Acctcd;
                //                    name = cvEntryTemp.AcctName;
                //                }

                //                if (acct != cvEntryTemp.Acctcd)
                //                {
                //                    acct = "SUM OF ";
                //                }
                //            }
                //        }

                //        if (cv.VatShouldBe >= amt - 5 && cv.VatShouldBe <= amt + 5)
                //        {
                //            if (acct == "SUM OF ")
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = "MULTIPLE";
                //                cv.VatDesc = $"SUM OF {ctr.ToString()} ENTRIES";
                //                cv.VatOk = true;
                //            }
                //            else
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = acct;
                //                cv.VatDesc = name;
                //                cv.VatOk = true;
                //            }
                //        }
                //    }

                //    // credit excluded (2) vat
                //    if (!cv.VatAmt.HasValue && cv.VatAmt == 0)
                //    {
                //        var amt = 0m;
                //        var ctr = 0m;
                //        var acct = string.Empty;
                //        var name = string.Empty;

                //        var acctList = new[] {
                //            "101060200",
                //            "101060300",
                //            "101010100"
                //        };

                //        var cvEntriesTemp = cvEntries2
                //            .Where(c => c.CvNo == cv.Header.CvNo
                //            && (acctList.Contains(c.Acctcd.Trim()) || c.Acctcd.Trim().StartsWith("201030"))
                //            && !c.Acctcd.Trim().StartsWith("V00")).ToList();

                //        if (cvEntriesTemp != null)
                //        {
                //            foreach (var cvEntryTemp in cvEntriesTemp)
                //            {
                //                amt = amt + (cvEntryTemp.DrCr ? 0 : cvEntryTemp.Amount) ?? 0m;
                //                ctr = ctr + 1;

                //                if (acct == string.Empty)
                //                {
                //                    acct = cvEntryTemp.Acctcd;
                //                    name = cvEntryTemp.AcctName;
                //                }

                //                if (acct != cvEntryTemp.Acctcd)
                //                {
                //                    acct = "SUM OF ";
                //                }
                //            }
                //        }

                //        if (cv.VatShouldBe >= amt - 5 && cv.VatShouldBe <= amt + 5)
                //        {
                //            if (acct == "SUM OF ")
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = "MULTIPLE";
                //                cv.VatDesc = $"SUM OF {ctr.ToString()} ENTRIES";
                //                cv.VatOk = true;
                //            }
                //            else
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = acct;
                //                cv.VatDesc = name;
                //                cv.VatOk = true;
                //            }
                //        }
                //    }
                //}

                //#endregion == EWT and VAT ==

                //#region == Trial Balances ==

                //// TB
                //var curtrialbal = cvEntries2
                //    .GroupBy(c => new
                //    {
                //        c.Acctcd,
                //        c.AcctName
                //    })
                //    .Select(c => new CurtrialbalViewModel
                //    {
                //        Acctcd = c.Key.Acctcd,
                //        AcctName = c.Key.AcctName,
                //        Debit = c.Sum(x => x.DrCr ? 0m : x.Amount),
                //        Credit = c.Sum(x => x.DrCr ? x.Amount : 0),
                //        Bal = 0m
                //    })
                //    .OrderBy(c => c.Acctcd)
                //    .ThenBy(c => c.AcctName)
                //    .ToList();

                //foreach(var trialBal in curtrialbal)
                //{
                //    trialBal.Bal = trialBal.Debit - trialBal.Credit;
                //}

                //// TB DOC
                //var curtrialbaldoc = cvEntries2
                //    .Where(c => c.BsNo.Contains("DOC") && !c.BsNo.Contains("UNDOC")
                //    && !c.BsNo.Contains("PUR") && !c.BsNo.Contains("HAU")
                //    && !c.BsNo.Contains("PAY"))
                //    .GroupBy(c => new
                //    {
                //        c.Acctcd,
                //        c.AcctName
                //    })
                //    .Select(c => new CurtrialbalViewModel
                //    {
                //        Acctcd = c.Key.Acctcd,
                //        AcctName = c.Key.AcctName,
                //        Debit = c.Sum(x => x.DrCr ? 0m : x.Amount),
                //        Credit = c.Sum(x => x.DrCr ? x.Amount : 0),
                //        Bal = 0m
                //    })
                //    .OrderBy(c => c.Acctcd)
                //    .ThenBy(c => c.AcctName)
                //    .ToList();

                //foreach (var trialBal in curtrialbaldoc)
                //{
                //    trialBal.Bal = trialBal.Debit - trialBal.Credit;
                //}

                //// TB UNDOC
                //var curtrialbalundoc = cvEntries2
                //    .Where(c => c.BsNo.Contains("UNDOC") && !c.BsNo.Contains("PUR") 
                //    && !c.BsNo.Contains("HAU") && !c.BsNo.Contains("PAY"))
                //    .GroupBy(c => new
                //    {
                //        c.Acctcd,
                //        c.AcctName
                //    })
                //    .Select(c => new CurtrialbalViewModel
                //    {
                //        Acctcd = c.Key.Acctcd,
                //        AcctName = c.Key.AcctName,
                //        Debit = c.Sum(x => x.DrCr ? 0m : x.Amount),
                //        Credit = c.Sum(x => x.DrCr ? x.Amount : 0),
                //        Bal = 0m
                //    })
                //    .OrderBy(c => c.Acctcd)
                //    .ThenBy(c => c.AcctName)
                //    .ToList();

                //foreach (var trialBal in curtrialbalundoc)
                //{
                //    trialBal.Bal = trialBal.Debit - trialBal.Credit;
                //}

                //// TB PURDOC
                //var curtrialbalpurdoc = cvEntries2
                //    .Where(c => c.BsNo.Contains("DOC") && !c.BsNo.Contains("UNDOC") 
                //    && c.BsNo.Contains("PUR"))
                //    .GroupBy(c => new
                //    {
                //        c.Acctcd,
                //        c.AcctName
                //    })
                //    .Select(c => new CurtrialbalViewModel
                //    {
                //        Acctcd = c.Key.Acctcd,
                //        AcctName = c.Key.AcctName,
                //        Debit = c.Sum(x => x.DrCr ? 0m : x.Amount),
                //        Credit = c.Sum(x => x.DrCr ? x.Amount : 0),
                //        Bal = 0m
                //    })
                //    .OrderBy(c => c.Acctcd)
                //    .ThenBy(c => c.AcctName)
                //    .ToList();

                //foreach (var trialBal in curtrialbalpurdoc)
                //{
                //    trialBal.Bal = trialBal.Debit - trialBal.Credit;
                //}

                //// TB PURUNDOC
                //var curtrialbalpurundoc = cvEntries2
                //    .Where(c => c.BsNo.Contains("UNDOC") && c.BsNo.Contains("PUR"))
                //    .GroupBy(c => new
                //    {
                //        c.Acctcd,
                //        c.AcctName
                //    })
                //    .Select(c => new CurtrialbalViewModel
                //    {
                //        Acctcd = c.Key.Acctcd,
                //        AcctName = c.Key.AcctName,
                //        Debit = c.Sum(x => x.DrCr ? 0m : x.Amount),
                //        Credit = c.Sum(x => x.DrCr ? x.Amount : 0),
                //        Bal = 0m
                //    })
                //    .OrderBy(c => c.Acctcd)
                //    .ThenBy(c => c.AcctName)
                //    .ToList();

                //foreach (var trialBal in curtrialbalpurundoc)
                //{
                //    trialBal.Bal = trialBal.Debit - trialBal.Credit;
                //}

                //// TB HAUDOC
                //var curtrialbalhaudoc = cvEntries2
                //    .Where(c => c.BsNo.Contains("DOC") 
                //    && c.BsNo.Contains("HAU")
                //    && !c.BsNo.Contains("UNDOC"))
                //    .GroupBy(c => new
                //    {
                //        c.Acctcd,
                //        c.AcctName
                //    })
                //    .Select(c => new CurtrialbalViewModel
                //    {
                //        Acctcd = c.Key.Acctcd,
                //        AcctName = c.Key.AcctName,
                //        Debit = c.Sum(x => x.DrCr ? 0m : x.Amount),
                //        Credit = c.Sum(x => x.DrCr ? x.Amount : 0),
                //        Bal = 0m
                //    })
                //    .OrderBy(c => c.Acctcd)
                //    .ThenBy(c => c.AcctName)
                //    .ToList();

                //foreach (var trialBal in curtrialbalhaudoc)
                //{
                //    trialBal.Bal = trialBal.Debit - trialBal.Credit;
                //}

                //// TB HAUUNDOC
                //var curtrialbalhauundoc = cvEntries2
                //    .Where(c => c.BsNo.Contains("UNDOC") 
                //    && c.BsNo.Contains("HAU"))
                //    .GroupBy(c => new
                //    {
                //        c.Acctcd,
                //        c.AcctName
                //    })
                //    .Select(c => new CurtrialbalViewModel
                //    {
                //        Acctcd = c.Key.Acctcd,
                //        AcctName = c.Key.AcctName,
                //        Debit = c.Sum(x => x.DrCr ? 0m : x.Amount),
                //        Credit = c.Sum(x => x.DrCr ? x.Amount : 0),
                //        Bal = 0m
                //    })
                //    .OrderBy(c => c.Acctcd)
                //    .ThenBy(c => c.AcctName)
                //    .ToList();

                //foreach (var trialBal in curtrialbalhauundoc)
                //{
                //    trialBal.Bal = trialBal.Debit - trialBal.Credit;
                //}

                //// TB PAYDOC
                //var curtrialbalpaydoc = cvEntries2
                //    .Where(c => c.BsNo.Contains("DOC") 
                //    && c.BsNo.Contains("PAY")
                //    && !c.BsNo.Contains("UNDOC"))
                //    .GroupBy(c => new
                //    {
                //        c.Acctcd,
                //        c.AcctName
                //    })
                //    .Select(c => new CurtrialbalViewModel
                //    {
                //        Acctcd = c.Key.Acctcd,
                //        AcctName = c.Key.AcctName,
                //        Debit = c.Sum(x => x.DrCr ? 0m : x.Amount),
                //        Credit = c.Sum(x => x.DrCr ? x.Amount : 0),
                //        Bal = 0m
                //    })
                //    .OrderBy(c => c.Acctcd)
                //    .ThenBy(c => c.AcctName)
                //    .ToList();

                //foreach (var trialBal in curtrialbalpaydoc)
                //{
                //    trialBal.Bal = trialBal.Debit - trialBal.Credit;
                //}

                //// TB PAYUNDOC
                //var curtrialbalpayundoc = cvEntries2
                //    .Where(c => c.BsNo.Contains("UNDOC") 
                //    && c.BsNo.Contains("PAY"))
                //    .GroupBy(c => new
                //    {
                //        c.Acctcd,
                //        c.AcctName
                //    })
                //    .Select(c => new CurtrialbalViewModel
                //    {
                //        Acctcd = c.Key.Acctcd,
                //        AcctName = c.Key.AcctName,
                //        Debit = c.Sum(x => x.DrCr ? 0m : x.Amount),
                //        Credit = c.Sum(x => x.DrCr ? x.Amount : 0),
                //        Bal = 0m
                //    })
                //    .OrderBy(c => c.Acctcd)
                //    .ThenBy(c => c.AcctName)
                //    .ToList();

                //foreach (var trialBal in curtrialbalpayundoc)
                //{
                //    trialBal.Bal = trialBal.Debit - trialBal.Credit;
                //}

                //TempData["success"] = "Success!";

                //#endregion == Trial Balances ==

                //#region == Final Cursors ==

                //var tempall = cvDetailx
                //    .Select(cv => cv.CvNo)
                //    .Distinct()
                //    .Order()
                //    .ToList();

                //var temp2all = cvDetailx
                //    .GroupBy(cv => cv.CvNo)
                //    .Select(cv => new
                //    {
                //        CvNo = cv.Key,
                //        Ctr = cv.Count(),
                //        Records = cv.ToList(),
                //        Header = cvHeaderx.Where(h => h.CvNo == cv.Key).FirstOrDefault(),
                //        Company = companies.Where(c => c.BankCode == cvHeaderx.Where(h => h.CvNo == cv.Key).FirstOrDefault().BankCode).FirstOrDefault()
                //    })
                //    .OrderByDescending(cv => cv.Ctr)
                //    .ToList();

                //var curcvheaderall = temp2all
                //    .Select(d => new CurcvheaderViewModel
                //    {
                //        CvDtl = d.CvNo,
                //        Header = d.Header,
                //        CashpoDate = new DateOnly(),
                //        DCRDate = new DateOnly(),
                //        Company = d.Company == null ? (string)null : d.Company.Co
                //    })
                //    .Where(d => d.Header.TranDate >= viewModel.DateFrom && d.Header.TranDate <= viewModel.DateTo)
                //    .ToList();

                //var cvDetailGroupAll = cvDetailx
                //    .GroupBy(d => new
                //    {
                //        d.CvNo,
                //        d.DrCr,
                //        d.Acctcd
                //    })
                //    .Select(g => new
                //    {
                //        g.Key.CvNo,
                //        g.Key.DrCr,
                //        g.Key.Acctcd,
                //        Amount = g.Sum(x => x.Amount),
                //        NEntry = g.Count(),
                //        SeqId = g.Max(x => x.SeqId)
                //    })
                //    .OrderBy(x => x.CvNo)
                //    .ThenBy(x => x.SeqId)
                //    .ToList();

                //var cvDetailGroup2All = cvDetailx
                //    .Select(d => new
                //    {
                //        d.CvNo,
                //        d.DrCr,
                //        d.Acctcd,
                //        d.Amount,
                //        d.SeqId,
                //        d.IsDisplayEntry
                //    })
                //    .OrderBy(d => d.CvNo)
                //    .ThenBy(d => d.SeqId)
                //    .ToList();

                //var detail_accountsall = cvDetailGroup
                //    .Select(d => new DetailAccountsViewModel
                //    {
                //        DrCr = d.DrCr,
                //        Acctcd = d.Acctcd,
                //        AcctName = chartOfAccounts.Where(coa => coa.AccountNumber == d.Acctcd).FirstOrDefault().AccountName,
                //        ColNum = (decimal)0
                //    })
                //    .Distinct()
                //    .OrderBy(d => d.Acctcd)
                //    .ToList();

                //var acctgroupall = detail_accounts
                //    .GroupBy(d => new
                //    {
                //        d.DrCr,
                //        d.Acctcd
                //    })
                //    .Select(d => new AcctGroupViewModel
                //    {
                //        Drcr = d.Key.DrCr,
                //        Acc = d.Key.Acctcd,
                //        Ctr = d.Count()
                //    })
                //    .OrderBy(d => d.Acc)
                //    .ToList();

                //var cventriesall = (from d in curcvheader
                //                 join h in cvDetailGroup on d.Header.CvNo equals h.CvNo
                //                 join a in detail_accountsall
                //                     on new { AcctCd = h.Acctcd.Trim(), DrCr = h.DrCr }
                //                     equals new { AcctCd = a.Acctcd.Trim(), DrCr = a.DrCr }
                //                     into leftJoin
                //                 from c in leftJoin.DefaultIfEmpty()
                //                 select new CvEntriesViewModel
                //                 {
                //                     CvNo = d.Header.CvNo,
                //                     TranDate = d.Header.TranDate,
                //                     Payee = d.Header.Payee,
                //                     BankCode = d.Header.BankCode,
                //                     BankName = d.Header.BankName,
                //                     CvAmount = d.Header.Amount,
                //                     CheckNo = d.Header.CheckNo,
                //                     CheckDate = d.Header.CheckDate,
                //                     BsNo = d.Header.BsNo,
                //                     ChkClear = d.Header.ChkClear,
                //                     Particulars = d.Header.Particulars,
                //                     Amount = h.Amount,
                //                     DrCr = h.DrCr,
                //                     Acctcd = h.Acctcd,
                //                     AcctName = c.AcctName,
                //                     ColNum = c.ColNum,
                //                     Reference = d.Header.Reference
                //                 })
                //    .OrderBy(h => h.CvNo)
                //    .ThenBy(h => h.Acctcd)
                //    .ToList();

                //// Verified
                //var cvEntries2all = (from a in curcvheaderall
                //                  from b in cvDetailGroup2All
                //                  where a.Header.CvNo == b.CvNo || a.Header.Reference == b.CvNo
                //                  join c in detail_accountsall
                //                      on new { AcctCd = b.Acctcd.Trim(), DrCr = b.DrCr }
                //                      equals new { AcctCd = c.Acctcd.Trim(), DrCr = c.DrCr }
                //                      into leftJoin
                //                  from c in leftJoin.DefaultIfEmpty()
                //                  select new CvEntriesViewModel
                //                  {
                //                      CvNo = a.Header.CvNo,
                //                      TranDate = a.Header.TranDate,
                //                      Payee = a.Header.Payee,
                //                      BankCode = a.Header.BankCode,
                //                      BankName = a.Header.BankName,
                //                      CvAmount = a.Header.Amount,
                //                      CheckNo = a.Header.CheckNo,
                //                      CheckDate = a.Header.CheckDate,
                //                      BsNo = a.Header.BsNo,
                //                      ChkClear = a.Header.ChkClear,
                //                      Particulars = a.Header.Particulars,
                //                      Amount = b.Amount,
                //                      DrCr = b.DrCr,
                //                      Acctcd = b.Acctcd,
                //                      AcctName = c != null ? c.AcctName : null,
                //                      ColNum = c != null ? c.ColNum : 0,
                //                      Reference = a.Header.Reference,
                //                      Seqid = b.SeqId,
                //                      CvType = a.Header.CvType,
                //                      Category = a.Header.Category,
                //                      IsDisplayEntry = b.IsDisplayEntry
                //                  })
                //    .OrderBy(x => x.CvNo)
                //    .ThenBy(x => x.ColNum)
                //    .ToList();

                //#endregion == Final Joins ==

                //#region == Assign DCR and 2307 values to cvall ==

                //// DCR
                //foreach (var cv in curcvheaderall)
                //{
                //    var dcrEntry = dcrMainDisbursements.Where(dcr => dcr.VoucherNo.Trim() == cv.Header.CvNo.Trim()).FirstOrDefault();

                //    if (dcrEntry != null)
                //    {
                //        cv.CashpoDate = dcrEntry.CashPoDate;
                //        cv.DCRDate = dcrEntry.DcrDate;
                //        cv.AccountNo = dcrEntry.AccountNo;
                //        cv.StnCode = dcrEntry.StnCode;
                //        if (dcrEntry.AccountNo != cv.Header.BankCode)
                //        {
                //            cv.Rem = "DIFFERENT BANK CODE";
                //        }
                //    }
                //    else
                //    {
                //        cv.CashpoDate = null;
                //        cv.DCRDate = null;
                //    }
                //}

                //// 2307
                //foreach (var cv in curcvheaderall)
                //{
                //    var cur = cur2307.Where(c => c.CvNo.Trim() == cv.Header.CvNo.Trim() && c.DownloadFrom.Contains("IBS")).FirstOrDefault();

                //    if (cur != null)
                //    {
                //        cv.DateFrom = cur.DateFrom;
                //        cv.DateTo = cur.DateTo;
                //        cv.AtcCode = cur.AtcCode;
                //        cv.AtcDesc = cur.Desc;
                //        cv.PayeeName = cur.PayeeName;
                //        cv.PayorName = cur.PayorName;
                //        cv.Tin = $"{cur.TinA}-{cur.TinB}-{cur.TinC}-{cur.TinD}";
                //        cv.Percent = cur.Percent;
                //        cv.Total = cur.Total;
                //        cv.Amt2307 = cur.Amount;
                //        cv.NMonth = cur.RMonth;
                //        cv.NYear = cur.RYear;
                //    }
                //}

                //#endregion == Assign DCR and 2307 values to cvall ==

                //#region == EWT and VAT to cvall ==

                //foreach (var cv in curcvheaderall)
                //{
                //    var vat = 0m;

                //    if ((cv.Amt2307.HasValue && cv.Amt2307 != 0) && (cv.Percent.HasValue && cv.Percent != 0))
                //    {
                //        cv.EwtShouldBe = cv.Amt2307 / (cv.Percent / 100);
                //    }

                //    var cvEntry = cventriesall
                //        .Where(c => c.CvNo == cv.Header.CvNo && c.Acctcd.Trim() == "101060200" && c.Acctcd.Trim() == "V00-17-101")
                //        .FirstOrDefault();

                //    // VAT
                //    if (cvEntry != null)
                //    {
                //        vat = (cvEntry.Amount ?? 0m) / 0.12m;

                //        cv.EwtAmt = cvEntry.Amount;
                //        cv.EwtAcctNo = cvEntry.Acctcd;
                //        cv.EwtDesc = cvEntry.AcctName;
                //        cv.EwtOk = true;
                //        cv.VatShouldBe = vat;

                //        var cvEntry2 = cventriesall
                //                .Where(c => c.CvNo.Trim() == cv.Header.CvNo.Trim() && c.Amount >= vat - 1 && c.Amount <= vat + 1)
                //                .FirstOrDefault();

                //        if (cvEntry2 != null)
                //        {
                //            cv.VatAmt = cvEntry2.Amount;
                //            cv.VatAcctNo = cvEntry2.Acctcd;
                //            cv.VatDesc = cvEntry2.AcctName;
                //            cv.VatOk = true;
                //        }
                //    }

                //    // credit included (1) ewt
                //    if (!cv.VatAmt.HasValue && cv.VatAmt == 0)
                //    {
                //        var amt = 0m;
                //        var ctr = 0m;
                //        var acct = string.Empty;
                //        var name = string.Empty;

                //        var acctList = new[] { "101060200", "101060300", "101010100" };

                //        var cvEntriesTemp = cvEntries2all
                //            .Where(c => c.CvNo == cv.Header.CvNo
                //            && (acctList.Contains(c.Acctcd.Trim()) || c.Acctcd.Trim().StartsWith("201030"))
                //            && !c.Acctcd.Trim().StartsWith("V00")).ToList();

                //        if (cvEntriesTemp != null)
                //        {
                //            foreach (var cvEntryTemp in cvEntriesTemp)
                //            {
                //                amt = amt + (cvEntryTemp.DrCr ? (cvEntryTemp.Amount * -1) : cvEntryTemp.Amount) ?? 0m;
                //                ctr = ctr + 1;

                //                if (acct == string.Empty || acct == null)
                //                {
                //                    acct = cvEntryTemp.Acctcd;
                //                    name = cvEntryTemp.AcctName;
                //                }

                //                if (acct != cvEntryTemp.Acctcd)
                //                {
                //                    acct = "SUM OF ";
                //                }
                //            }
                //        }

                //        if (cv.EwtShouldBe >= amt - 5 && cv.EwtShouldBe <= amt + 5)
                //        {
                //            if (acct == "SUM OF ")
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = "MULTIPLE";
                //                cv.VatDesc = $"SUM OF {ctr.ToString()} ENTRIES";
                //                cv.VatOk = true;
                //            }
                //            else
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = acct;
                //                cv.VatDesc = name;
                //                cv.VatOk = true;
                //            }
                //        }
                //    }

                //    // credit excluded (1) ewt
                //    if (!cv.VatAmt.HasValue && cv.VatAmt == 0)
                //    {
                //        var amt = 0m;
                //        var ctr = 0m;
                //        var acct = string.Empty;
                //        var name = string.Empty;

                //        var acctList = new[] { "101060200", "101060300", "101010100" };

                //        var cvEntriesTemp = cvEntries2all
                //            .Where(c => c.CvNo == cv.Header.CvNo
                //            && (acctList.Contains(c.Acctcd.Trim()) || c.Acctcd.Trim().StartsWith("201030"))
                //            && !c.Acctcd.Trim().StartsWith("V00")).ToList();

                //        if (cvEntriesTemp != null)
                //        {
                //            foreach (var cvEntryTemp in cvEntriesTemp)
                //            {
                //                amt = amt + (cvEntryTemp.DrCr ? 0 : cvEntryTemp.Amount) ?? 0m;
                //                ctr = ctr + 1;

                //                if (acct == string.Empty)
                //                {
                //                    acct = cvEntryTemp.Acctcd;
                //                    name = cvEntryTemp.AcctName;
                //                }

                //                if (acct != cvEntryTemp.Acctcd)
                //                {
                //                    acct = "SUM OF ";
                //                }
                //            }
                //        }

                //        if (cv.EwtShouldBe >= amt - 5 && cv.EwtShouldBe <= amt + 5)
                //        {
                //            if (acct == "SUM OF ")
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = "MULTIPLE";
                //                cv.VatDesc = $"SUM OF {ctr.ToString()} ENTRIES";
                //                cv.VatOk = true;
                //            }
                //            else
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = acct;
                //                cv.VatDesc = name;
                //                cv.VatOk = true;
                //            }
                //        }
                //    }

                //    // credit included (2) vat
                //    if (!cv.VatAmt.HasValue && cv.VatAmt == 0)
                //    {
                //        var amt = 0m;
                //        var ctr = 0m;
                //        var acct = string.Empty;
                //        var name = string.Empty;

                //        var acctList = new[] { "101060200", "101060300", "101010100" };

                //        var cvEntriesTemp = cvEntries2all
                //            .Where(c => c.CvNo == cv.Header.CvNo
                //            && (acctList.Contains(c.Acctcd.Trim()) || c.Acctcd.Trim().StartsWith("201030"))
                //            && !c.Acctcd.Trim().StartsWith("V00")).ToList();

                //        if (cvEntriesTemp != null)
                //        {
                //            foreach (var cvEntryTemp in cvEntriesTemp)
                //            {
                //                amt = amt + (cvEntryTemp.DrCr ? cvEntryTemp.Amount * -1 : cvEntryTemp.Amount) ?? 0m;
                //                ctr = ctr + 1;

                //                if (acct == string.Empty)
                //                {
                //                    acct = cvEntryTemp.Acctcd;
                //                    name = cvEntryTemp.AcctName;
                //                }

                //                if (acct != cvEntryTemp.Acctcd)
                //                {
                //                    acct = "SUM OF ";
                //                }
                //            }
                //        }

                //        if (cv.VatShouldBe >= amt - 5 && cv.VatShouldBe <= amt + 5)
                //        {
                //            if (acct == "SUM OF ")
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = "MULTIPLE";
                //                cv.VatDesc = $"SUM OF {ctr.ToString()} ENTRIES";
                //                cv.VatOk = true;
                //            }
                //            else
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = acct;
                //                cv.VatDesc = name;
                //                cv.VatOk = true;
                //            }
                //        }
                //    }

                //    // credit excluded (2) vat
                //    if (!cv.VatAmt.HasValue && cv.VatAmt == 0)
                //    {
                //        var amt = 0m;
                //        var ctr = 0m;
                //        var acct = string.Empty;
                //        var name = string.Empty;

                //        var acctList = new[] { "101060200", "101060300", "101010100" };

                //        var cvEntriesTemp = cvEntries2all
                //            .Where(c => c.CvNo == cv.Header.CvNo
                //            && (acctList.Contains(c.Acctcd.Trim()) || c.Acctcd.Trim().StartsWith("201030"))
                //            && !c.Acctcd.Trim().StartsWith("V00")).ToList();

                //        if (cvEntriesTemp != null)
                //        {
                //            foreach (var cvEntryTemp in cvEntriesTemp)
                //            {
                //                amt = amt + (cvEntryTemp.DrCr ? 0 : cvEntryTemp.Amount) ?? 0m;
                //                ctr = ctr + 1;

                //                if (acct == string.Empty)
                //                {
                //                    acct = cvEntryTemp.Acctcd;
                //                    name = cvEntryTemp.AcctName;
                //                }

                //                if (acct != cvEntryTemp.Acctcd)
                //                {
                //                    acct = "SUM OF ";
                //                }
                //            }
                //        }

                //        if (cv.VatShouldBe >= amt - 5 && cv.VatShouldBe <= amt + 5)
                //        {
                //            if (acct == "SUM OF ")
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = "MULTIPLE";
                //                cv.VatDesc = $"SUM OF {ctr.ToString()} ENTRIES";
                //                cv.VatOk = true;
                //            }
                //            else
                //            {
                //                cv.VatAmt = amt;
                //                cv.VatAcctNo = acct;
                //                cv.VatDesc = name;
                //                cv.VatOk = true;
                //            }
                //        }
                //    }
                //}

                //#endregion == EWT and VAT cvall ==

                //#region == Delete Invoicing ==

                //curcvheaderall = curcvheaderall.Where(cv => cv.Header.CvType.Trim() != "Invoicing").ToList();
                //cventriesall = cventriesall.Where(cv => cv.CvType.Trim() != "Invoicing").ToList();
                //cvEntries2all = cvEntries2all.Where(cv => cv.CvType.Trim() != "Invoicing").ToList();

                //#endregion == Delete Invoicing

                //#region == Trial Balance all cv ==

                //// TB all cv
                //var curtrialbal2all = cvEntries2all
                //    .GroupBy(c => new
                //    {
                //        c.Acctcd,
                //        c.AcctName
                //    })
                //    .Select(c => new CurtrialbalViewModel
                //    {
                //        Acctcd = c.Key.Acctcd,
                //        AcctName = c.Key.AcctName,
                //        Debit = c.Sum(x => x.DrCr ? 0m : x.Amount),
                //        Credit = c.Sum(x => x.DrCr ? x.Amount : 0),
                //        Bal = 0m
                //    })
                //    .OrderBy(c => c.Acctcd)
                //    .ThenBy(c => c.AcctName)
                //    .ToList();

                //foreach (var trialBal in curtrialbal2all)
                //{
                //    trialBal.Bal = trialBal.Debit - trialBal.Credit;
                //}

                //#endregion == Trial balance all cv ==

                //#region == Assign DCR and 2307 values to cvall ==

                //// DCR
                //foreach (var cv in cvEntries2all)
                //{
                //    var cvno = cv.CvNo.Trim();
                //    var reference = cv.Reference.Trim();
                //    var name = cv.AcctName.Trim();
                //    var bank = cv.BankCode.Trim();

                //    if (cv.Acctcd.StartsWith("2010302"))
                //    {
                //        var selected2307 = cur2307
                //            .Where(cur => (cur.CvNo == cvno || cur.CvNo == reference) && cur.Description == name)
                //            .FirstOrDefault();

                //        if (selected2307 != null)
                //        {
                //            cv.DateFrom = selected2307.DateFrom;
                //            cv.DateTo = selected2307.DateTo;
                //            cv.PayorName = selected2307.PayorName;
                //            cv.PayeeName = selected2307.PayeeName;
                //            cv.Tin = $"{selected2307.TinA}-{selected2307.TinB}-{selected2307.TinC}-{selected2307.TinD}";
                //            cv.AtcCode = selected2307.AtcCode;
                //            cv.Desc = selected2307.Desc;
                //            cv.Percent = selected2307.Percent;
                //            cv.NTotal = selected2307.Total;
                //            cv.NAmount = selected2307.Amount;
                //            cv.NMonth = selected2307.RMonth;
                //            cv.NYear = selected2307.RYear;
                //        }
                //    }

                //    if (cv.Acctcd.Trim() == "101010100")
                //    {
                //        var selectedDcr = dcrMainDisbursements
                //            .Where(cur => cur.AccountNo == bank && cur.VoucherNo.Trim() == cvno)
                //            .FirstOrDefault();

                //        if (selectedDcr!= null)
                //        {
                //            cv.AccountNo = selectedDcr.AccountNo;
                //            cv.CashpoDate = selectedDcr.CashPoDate;
                //            cv.DcrDate = selectedDcr.DcrDate;
                //        }
                //    }
                //}

                //#endregion == Assign DCR and 2307 values to cvall ==

                #endregion == Data related ==

                #region == Generate Report ==

                if (viewModel.FileFormat == "Excel")
                {
                    var dateFrom = viewModel.DateFrom;
                    var dateTo = viewModel.DateTo;

                    using var package = new ExcelPackage();

                    var report1 = package.Workbook.Worksheets.Add("Rpt#1");
                    _excel.ProcessReport1(report1, viewModel);









                    // var fileName = $"Purchase_Order_Report_{DateTimeHelper.GetCurrentPhilippineTime():yyyyddMMHHmmss}.xlsx";
                    var fileName = $"Tax_Report.xlsx";
                    var stream = new MemoryStream();
                    await package.SaveAsAsync(stream, cancellationToken);
                    stream.Position = 0;
                    return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
                else
                {
                    // PDF Code
                }

                #endregion == Generate Report ==
            }
            catch (Exception ex)
            {
                TempData["error"] = ex.Message;
            }

            return RedirectToAction(nameof(Index));
        }
    }
}
