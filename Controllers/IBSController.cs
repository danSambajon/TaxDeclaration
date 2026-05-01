using Microsoft.AspNetCore.Mvc;

namespace TaxDeclaration.Controllers
{
    public class IBSController : Controller
    {
        // go to data folder(server)
        // copy the station, fastcv, fastledger, fastcoa, stationx from station(with recpos), taxdeclare, company(with recpos)
        // interface contains:
        // date range
        // fast (companies: mmsi, mobility, syvill)
        // arnontrade (companies: bienes, mcy, mnv personal)
        // ibs: disbursement, ibs hauling, ibs sales, ibs purchases (company: filpride only)
        // options excel or pdf

        public IActionResult Index()
        {
            return View();
        }
    }
}
