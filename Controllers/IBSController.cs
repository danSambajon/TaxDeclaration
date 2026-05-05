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
            var stations = _dbfService.GetStationsFromDbf();
            var fastCvs = _dbfService.GetFastCvsFromDbf();
            var companies = _dbfService.GetCompaniesFromDbf();

            var checkVoucherHeaders = new List<FilprideCheckVoucherHeader>();
            var checkVoucherDetails = new List<FilprideCheckVoucherDetail>();
            var chartOfAccounts = new List<FilprideChartOfAccount>();

            try
            {
                //stations = _dbfService.GetStationsFromDbf();
                //fastCvs = _dbfService.GetFastCvsFromDbf();
                //companies = _dbfService.GetCompaniesFromDbf();

                checkVoucherHeaders = await _ibsConnectionService.GetCheckVoucherHeaders();
                checkVoucherDetails = await _ibsConnectionService.GetCheckVoucherDetails(checkVoucherHeaders.Select(h => h.CheckVoucherHeaderNo).ToList());
                //chartOfAccounts = await _ibsConnectionService.GetChartOfAccounts();
            }
            catch (Exception ex)
            {
                return Content(ex.Message);
            }

            return RedirectToAction(nameof(Index));
        }
    }
}
