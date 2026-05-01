using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using TaxDeclaration.Models.ViewModels;

namespace TaxDeclaration.Controllers
{
    public class ARNTradeController : Controller
    {
        public IActionResult Index()
        {
            var viewModel = new GenerateTaxDeclarationViewModel
            {
                Company = "MMSI",
                CompanyChoices = new List<SelectListItem>
                {
                    new SelectListItem { Text = "BIENES", Value = "BIENES" },
                    new SelectListItem { Text = "MCY", Value = "MCY" },
                    new SelectListItem { Text = "MNV PERSONAL", Value = "MNV PERSONAL" }
                }
            };

            viewModel.SelectedCompany = viewModel.CompanyChoices.FirstOrDefault()!.Value;

            return View(viewModel);
        }
    }
}
