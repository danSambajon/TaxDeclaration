using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using TaxDeclaration.Models.ViewModels;
using TaxDeclaration.Services.Dbf;

namespace TaxDeclaration.Controllers
{
    public class IBSController : Controller
    {
        private readonly DbfService _dbfService;

        public IBSController(DbfService dbfService)
        {
            _dbfService = dbfService;
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
            try
            {
                var stations = _dbfService.GetStationsFromDbf();
            }
            catch (Exception ex)
            {
                return Content(ex.Message);
            }

            return RedirectToAction(nameof(Index));
        }
    }
}
