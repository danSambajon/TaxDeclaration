using DbfDataReader;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using System.Threading;
using TaxDeclaration.Data;
using TaxDeclaration.Models;
using TaxDeclaration.Models.ViewModels;
using static System.Collections.Specialized.BitVector32;
using TaxDeclaration.Services.Dbf;

namespace TaxDeclaration.Controllers
{
    public class IBSController : Controller
    {
        private readonly ApplicationDbContext _dbContext;
        private readonly DbfService _dbfService;

        public IBSController(ApplicationDbContext dbContext, DbfService dbfService)
        {
            _dbContext = dbContext;
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
