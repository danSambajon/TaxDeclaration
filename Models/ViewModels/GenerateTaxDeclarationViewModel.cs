using System.ComponentModel.DataAnnotations;

namespace TaxDeclaration.Models.ViewModels
{
    public class GenerateTaxDeclarationViewModel
    {
        public DateOnly DateFrom { get; set; }
        public DateTime DateTo { get; set; }
        public required string Company { get; set; }
        public required string FileFormat { get; set; }
    }
}
