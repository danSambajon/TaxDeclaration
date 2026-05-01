using Microsoft.AspNetCore.Mvc;

namespace TaxDeclaration.Controllers
{
    public class IBSController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
