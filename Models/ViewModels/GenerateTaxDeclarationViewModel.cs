using Microsoft.AspNetCore.Mvc.Rendering;
using System.ComponentModel.DataAnnotations;

namespace TaxDeclaration.Models.ViewModels
{
    public class GenerateTaxDeclarationViewModel
    {
        public DateOnly DateFrom { get; set; }
        public DateTime DateTo { get; set; }
        public string Company { get; set; }
        public string FileFormat { get; set; }
        public string? IBSCategory { get; set; }
        public string? SelectedCompany { get; set; }
        public List<SelectListItem>? CompanyChoices { get; set; }
    }
}
