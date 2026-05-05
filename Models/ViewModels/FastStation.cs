namespace TaxDeclaration.Models.ViewModels
{
    public class FastStation
    {
        public string? Company {  get; set; }
        public string? StnCode { get; set; }
        public string? StnName { get; set; }
        public string? CodeName { get; set; }
        public string? DataFolder { get; set; }
        public bool Active { get; set; }
        public string? Sort { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string? CreatedBy { get; set; }
        public DateTime? EditedDate { get; set; }
        public string? EditedBy { get; set; }
    }
}
