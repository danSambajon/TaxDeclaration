using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using TaxDeclaration.Models.ViewModels;

namespace TaxDeclaration.Controllers
{
    public class FastController : Controller
    {
        public IActionResult Index()
        {
            var viewModel = new GenerateTaxDeclarationViewModel
            {
                Company = "MMSI",
                CompanyChoices = new List<SelectListItem>
                {
                    new SelectListItem { Text = "MMSI", Value = "MMSI" },
                    new SelectListItem { Text = "MOBILITY", Value = "MOBILITY" },
                    new SelectListItem { Text = "SYVILL", Value = "SYVILL" }
                }
            };

            viewModel.SelectedCompany = viewModel.CompanyChoices.FirstOrDefault()!.Value;

            return View(viewModel);
        }
    }
}
