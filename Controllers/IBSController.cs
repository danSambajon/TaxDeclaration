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
                stations = _dbfService.GetStationsFromDbf();
                fastCvs = _dbfService.GetFastCvsFromDbf();
                companies = _dbfService.GetCompaniesFromDbf();
                //dcrMainDisbursements = _dbfService.GetDCRMainDisbursementsFromDbf();

                // Get data from IBS
                checkVoucherHeaders = await _ibsConnectionService.GetCheckVoucherHeaders(viewModel.DateFrom, viewModel.DateTo);
                if (checkVoucherHeaders.Count == 0)
                {
                    throw new NullReferenceException("No check voucher headers found for the specified date range.");
                }

                checkVoucherDetails = await _ibsConnectionService.GetCheckVoucherDetails(checkVoucherHeaders.Select(h => h.CheckVoucherHeaderNo).ToList());
                chartOfAccounts = await _ibsConnectionService.GetChartOfAccounts();

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
