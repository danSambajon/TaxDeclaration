using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using TaxDeclaration.Models.ViewModels;

namespace TaxDeclaration.Controllers
{
    public class FastController : Controller
    {
        // go to data folder(server)
        // copy the station, fastcv, fastledger, fastcoa, stationx from station(with recpos), taxdeclare, company(with recpos)
        // interface contains:
        // date range DONE
        // fast (companies: mmsi, mobility, syvill)
        // arnontrade (companies: bienes, mcy, mnv personal)
        // ibs: disbursement, ibs hauling, ibs sales, ibs purchases (company: filpride only)
        // options excel or pdf

        public IActionResult Index()
        {
            var viewModel = new GenerateTaxDeclarationViewModel
            {
                Company = "MMSI",
                CompanyChoices = new List<SelectListItem>
                {
                    new SelectListItem { Text = "MMSI", Value = "MMSI" },
                    new SelectListItem { Text = "Mobility", Value = "Mobility" },
                    new SelectListItem { Text = "MNV Personal", Value = "MNV Personal" }
                }
            };

            viewModel.SelectedCompany = viewModel.CompanyChoices.FirstOrDefault()!.Value;

            return View(viewModel);
        }
    }
}
