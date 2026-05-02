using TaxDeclaration.Models;
using TaxDeclaration.Models.ViewModels;

namespace TaxDeclaration.Services.Dbf
{
    public class DbfService
    {
        public List<Station> GetStationsFromDbf()
        {
            var stations = new List<Station>();
            using var dbf = new DbfDataReader.DbfDataReader(DbfPaths.StationPath);

            while (dbf.Read())
            {
                var station = new Station
                {
                    Company = dbf["COMPANY"]?.ToString(),
                    StnCode = dbf["STNCODE"]?.ToString(),
                    StnName = dbf["STNNAME"]?.ToString(),
                    CodeName = dbf["CODENAME"]?.ToString(),
                    DataFolder = dbf["DATAFOLDER"]?.ToString(),
                    Active = dbf["ACTIVE"].ToString() == "True",
                    Sort = dbf["SORT"]?.ToString(),
                    CreatedDate = DateTime.TryParse(
                        dbf["CREATEDDAT"]?.ToString(),
                        out var createdDate
                        ) ? createdDate : null,
                    CreatedBy = dbf["CREATEDBY"]?.ToString(),
                    EditedDate = DateTime.TryParse(
                        dbf["EDITEDDATE"]?.ToString(),
                        out var editedDate
                        ) ? editedDate : null,
                    EditedBy = dbf["EDITEDBY"]?.ToString()
                };

                stations.Add(station);
            }

            return stations;
        }
    }
}
