using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using TaxDeclaration.Models;
using TaxDeclaration.Models.ViewModels;
using TaxDeclaration.Services;
using TaxDeclaration.Services.Dbf;

namespace TaxDeclaration.Controllers
{
    public class IBSController : Controller
    {
        private readonly DbfService _dbfService;
        private readonly IBSConnectionService _ibsConnectionService;

        public IBSController(DbfService dbfService, IBSConnectionService ibsConnectionService)
        {
            _dbfService = dbfService;
            _ibsConnectionService = ibsConnectionService;
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
            #region == Initialize Data Containers ==

            var stations = new List<FastStationViewModel>();
            var fastCvs = new List<FastCvViewModel>();
            var companies = new List<TaxDeclareCompanyViewModel>();
            var dcrMainDisbursements = new List<DCRMainDisbursementViewModel>();

            var checkVoucherHeaders = new List<FilprideCheckVoucherHeader>();
            var checkVoucherDetails = new List<FilprideCheckVoucherDetail>();
            var chartOfAccounts = new List<FilprideChartOfAccount>();

            #endregion == Initialize Data Containers ==


            try
            {
                // Get data from DBF files
                //stations = _dbfService.GetStationsFromDbf();
                //fastCvs = _dbfService.GetFastCvsFromDbf();
                companies = _dbfService.GetCompaniesFromDbf();
                //dcrMainDisbursements = _dbfService.GetDCRMainDisbursementsFromDbf();

                // 

                // Get data from IBS
                checkVoucherHeaders = await _ibsConnectionService
                    .GetCheckVoucherHeaders(viewModel.DateFrom, viewModel.DateTo);

                if (checkVoucherHeaders.Count == 0)
                {
                    throw new NullReferenceException("No check voucher headers found for the specified date range.");
                }

                checkVoucherDetails = await _ibsConnectionService
                    .GetCheckVoucherDetails(checkVoucherHeaders.Select(h => h.CheckVoucherHeaderNo).ToList(), checkVoucherHeaders.Select(h => h.Reference).ToList());

                chartOfAccounts = await _ibsConnectionService
                    .GetChartOfAccounts();

                

                // ASSIGN DETAILS TO HEADERS
                var cvHeadersWithDetails = new List<FilprideCheckVoucherHeader>();

                foreach (FilprideCheckVoucherHeader header in checkVoucherHeaders)
                {
                    var selectedHeader = header;
                    selectedHeader.Details = checkVoucherDetails.Where(d => d.TransactionNo == header.CheckVoucherHeaderNo).ToList();
                    selectedHeader.Referenced = checkVoucherDetails.Where(d => d.TransactionNo == header.Reference).ToList();
                    cvHeadersWithDetails.Add(selectedHeader);
                }



                // ASSIGN HEADER TO DETAILS(transno)
                var cvDetailsWithHeader = new List<FilprideCheckVoucherDetail>();
                var cvDetailsWithoutHeader = new List<FilprideCheckVoucherDetail>();

                foreach (FilprideCheckVoucherDetail detail in checkVoucherDetails)
                {
                    var selectedDetail = detail;
                    selectedDetail.Header = checkVoucherHeaders.Where(h => h.CheckVoucherHeaderNo == detail.TransactionNo).FirstOrDefault();

                    if (selectedDetail.Header == null)
                    {
                        cvDetailsWithoutHeader.Add(selectedDetail);
                    }
                    else
                    {
                        selectedDetail.CompanyVm = companies.Where(c => c.BankCode == selectedDetail.Header.BankAccountNumber).FirstOrDefault();
                        cvDetailsWithHeader.Add(selectedDetail);
                    }
                }



                // CHECKING VALUES
                var headersWithNoDetails = cvHeadersWithDetails.Where(h => h.Details.Count == 0).ToList();
                var headerWithDetails = checkVoucherHeaders.Where(h => h.Details.Count > 0).ToList();
                var detailsRegistered = checkVoucherHeaders.Sum(h => h.Details.Count);
                var headerCount = checkVoucherHeaders.Count;



                var cvHeaderx = checkVoucherHeaders
                    .Select(h => new
                    {
                        CvNo = h.CheckVoucherHeaderNo,
                        TranDate = h.Date,
                        h.Payee,
                        BankCode = h.BankAccountNumber,
                        BankName = h.BankAccountNumber,
                        h.CheckNo,
                        Amount = h.Total,
                        h.CheckDate,
                        h.Particulars,
                        h.CreatedBy,
                        AcctCd = (string)null,
                        BsNo = h.Type.ToUpper(),
                        CheckClearing = h.DcrDate,
                        NameCategory = (string)null,
                        IsCancelled = h.CanceledBy != null,
                        ChkClear = h.DcrDate,
                        h.Type,
                        h.Reference,
                        h.Category,
                        h.CvType
                    })
                    .ToList();



                var cvNoFromcvheaderx = cvHeaderx.Select(h => h.CvNo).ToList();
                var referenceFromcvheaderx = cvHeaderx.Select(h => h.Reference).ToList();

                var cvDetailx = checkVoucherDetails
                    .Where(d => !d.IsDisplayEntry 
                    && (cvNoFromcvheaderx.Contains(d.TransactionNo!.Trim()) || referenceFromcvheaderx.Contains(d.TransactionNo!.Trim())))
                    .Select(d => new
                    {
                        CvNo = d.TransactionNo,
                        SeqId = d.CheckVoucherDetailId,
                        Amount = d.Debit != 0 || d.Debit == null ? d.Debit : d.Credit,
                        DrCr = d.Debit == 0 || d.Debit == null,
                        Acctcd = d.AccountNo,
                        d.IsDisplayEntry,
                        d.Header
                    })
                        .ToList();


                var bankCodesOfCvDetailx = cvDetailx.Select(d => d.Header.BankAccountNumber).Distinct().ToList();


                // DETAILS: VAT INPUT AND GOVERNMENT PAYABLES 
                var temp = cvDetailx
                    .Where(d => d.Acctcd.StartsWith("101060200") || d.Acctcd.StartsWith("201030"))
                    .OrderBy(d => d.CvNo)
                    .Select(d => d.CvNo)
                    .Distinct()
                    .ToList();


                // DETAILS(GROUPED): VAT INPUT AND GOVERNMENT PAYABLES
                var temp2 = cvDetailx
                    .Where(d => d.Acctcd.StartsWith("101060200") || d.Acctcd.StartsWith("201030"))
                    .GroupBy(d => d.CvNo)
                    .Select(g => new
                    {
                        CvNo = g.Key,
                        Ctr = g.Count(),
                        Records = g.ToList(),
                        Header = cvHeaderx.Where(h => h.CvNo == g.Key).FirstOrDefault(),
                        Company = companies.Where(c => c.BankCode == cvHeaderx.Where(h => h.CvNo == g.Key).FirstOrDefault().BankCode).FirstOrDefault()
                    })
                    .OrderByDescending(g => g.Ctr)
                    .ToList();


                // DETAILS + HEADER: WITH EMPTY FIELDS
                var curcvheader = temp2
                    .Select(d => new
                    {
                        CvDtl = d.CvNo,
                        d.Header,
                        CashpoDate = new DateOnly(),
                        DCRDate = new DateOnly(),
                        AccountNo = (string)null,
                        StnCode = (string)null,
                        Company = d.Company == null ? (string)null : d.Company.Co,
                        Rem = (string)null,
                        DateFrom = new DateOnly(),
                        DateTo = new DateOnly(),
                        AtcCode = (string)null,
                        Percent = (decimal)0,
                        Payprname = (string)null,
                        PayeeName = (string)null,
                        Tin = (string)null,
                        Total = (decimal)0,
                        Amt2307 = (decimal)0,
                        NMonth = (decimal)0,
                        NYear = (decimal)0,
                        VatAmt = (decimal)0,
                        VatAcctNo = (string)null,
                        VatDesc = (string)null,
                        EwtAmt = (decimal)0,
                        EwtAcctNo = (string)null,
                        EwtDesc = (string)null,
                        VatShouldBe = (decimal)0,
                        VatVariance = (decimal)0,
                        VatOk = false,
                        EwtShouldBe = (decimal)0,
                        EwtVariance = (decimal)0,
                        EwtOk = false,
                        LedgerDebit = (decimal)0,
                        LedgerAcctNo = (string)null,
                        LedgerDesc = (string)null,
                        RowNum = (string)null
                    })
                    .ToList();


                var cvDetailGroup = cvDetailx
                    .Where(d => temp.Contains(d.CvNo))
                    .GroupBy(d => new
                    {
                        d.CvNo,
                        d.DrCr,
                        d.Acctcd
                    })
                    .Select(g => new
                    {
                        g.Key.CvNo,
                        g.Key.DrCr,
                        g.Key.Acctcd,
                        Amount = g.Sum(x => x.Amount),
                        NEntry = g.Count(),
                        SeqId = g.Max(x => x.SeqId)
                    })
                    .OrderBy(x => x.CvNo)
                    .ThenBy(x => x.SeqId)
                    .ToList();


                var cvDetailGroup2 = cvDetailx
                    .Where(d => temp.Contains(d.CvNo))
                    .OrderBy(d =>  d.CvNo )
                    .ThenBy(d => d.SeqId)
                    .Select(d => new
                    {
                        d.CvNo,
                        d.DrCr,
                        d.Acctcd,
                        AcctName = chartOfAccounts.Where(coa => coa.AccountNumber == d.Acctcd).FirstOrDefault().AccountName,
                        d.Amount,
                        d.SeqId
                    })
                    .ToList();


                var detail_accounts = cvDetailGroup
                    .Select(d => new
                    {
                        d.DrCr,
                        d.Acctcd,
                        Coa = chartOfAccounts.Where(coa => coa.AccountNumber == d.Acctcd).FirstOrDefault().AccountName,
                        ColNum = (decimal)0
                    })
                    .Distinct()
                    .OrderBy(d => d.Acctcd)
                    .ToList();


                var acctgroup = detail_accounts
                    .GroupBy(d => new
                    {
                        d.DrCr,
                        d.Acctcd
                    })
                    .Select(d => new
                    {
                        d.Key.DrCr,
                        Acc = d.Key.Acctcd,
                        Ctr = d.Count(),
                        Colnum = (decimal)0
                    })
                    .ToList();









                foreach (FilprideCheckVoucherDetail detail in checkVoucherDetails)
                {
                    var headers = checkVoucherHeaders.Where(h => h.CheckVoucherHeaderNo == detail.TransactionNo || h.Reference == detail.TransactionNo).ToList();



                    if (headers.Where(h => h.CvType == "Payment").FirstOrDefault() != null)
                    {

                    }

                    detail.Header = headers.FirstOrDefault();
                }

                TempData["success"] = "Success!";
            }
            catch (Exception ex)
            {
                TempData["error"] = ex.Message;
            }

            return RedirectToAction(nameof(Index));
        }
    }
}
