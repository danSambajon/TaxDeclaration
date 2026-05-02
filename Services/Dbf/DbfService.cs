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

        public List<FastCv> GetFastCvsFromDbf()
        {
            var fastCvs = new List<FastCv>();
            using var dbf = new DbfDataReader.DbfDataReader(DbfPaths.FastCvPath);
            var readStatus = dbf.Read();

            while (dbf.Read())
            {
                var fastCv = new FastCv
                {
                    Station = dbf["STATION"]?.ToString(),
                    Trans_no = dbf["TRANS_NO"]?.ToString(),
                    Voucher_no = dbf["VOUCHER_NO"]?.ToString(),
                    Vch_Date = DateOnly.TryParse(dbf["VCH_DATE"]?.ToString(), out var vchDate) ? vchDate : (DateOnly?)null,
                    Payee = dbf["PAYEE"]?.ToString(),
                    Amount = decimal.TryParse(dbf["AMOUNT"]?.ToString(), out var amount) ? amount : (decimal?)null,
                    PayFor1 = dbf["PAYFOR1"]?.ToString(),
                    Particular = dbf["PARTICULAR"]?.ToString(),
                    Qty = decimal.TryParse(dbf["QTY"]?.ToString(), out var qty) ? qty : (decimal?)null,
                    Unit_price = decimal.TryParse(dbf["UNIT_PRICE"]?.ToString(), out var unitPrice) ? unitPrice : (decimal?)null,
                    Vat = decimal.TryParse(dbf["VAT"]?.ToString(), out var vat) ? vat : (decimal?)null,
                    WTax = decimal.TryParse(dbf["WTAX"]?.ToString(), out var wTax) ? wTax : (decimal?)null,
                    Total = decimal.TryParse(dbf["TOTAL"]?.ToString(), out var total) ? total : (decimal?)null,
                    Ref1 = dbf["REF1"]?.ToString(),
                    Ref2 = dbf["REF2"]?.ToString(),
                    Red3 = dbf["RED3"]?.ToString(),
                    Others = dbf["OTHERS"]?.ToString(),
                    CheckNo = dbf["CHECKNO"]?.ToString(),
                    ChkDate = DateOnly.TryParse(dbf["CHKDATE"]?.ToString(), out var chkDate) ? chkDate : (DateOnly?)null,
                    Bank = dbf["BANK"]?.ToString(),
                    Acct_no = dbf["ACCT_NO"]?.ToString(),
                    Debit = decimal.TryParse(dbf["DEBIT"]?.ToString(), out var debit) ? debit : (decimal?)null,
                    Credit = decimal.TryParse(dbf["CREDIT"]?.ToString(), out var credit) ? credit : (decimal?)null,
                    Dcp_date = DateOnly.TryParse(dbf["DCP_DATE"]?.ToString(), out var dcpDate) ? dcpDate : (DateOnly?)null,
                    Dcr_date = DateOnly.TryParse(dbf["DCR_DATE"]?.ToString(), out var dcrDate) ? dcrDate : (DateOnly?)null,
                    Dte_posted = DateOnly.TryParse(dbf["DTE_POSTED"]?.ToString(), out var dtePosted) ? dtePosted : (DateOnly?)null,
                    Cancel = dbf["CANCEL"]?.ToString() == "TRUE",
                    Pre_by = dbf["PRE_BY"]?.ToString(),
                    UserID = dbf["USERID"]?.ToString(),
                    DtDownload = DateOnly.TryParse(dbf["DTDOWNLOAD"]?.ToString(), out var dtDownload) ? dtDownload : (DateOnly?)null,
                    Nawala = dbf["NAWALA"]?.ToString() == "TRUE",
                    Recnum = decimal.TryParse(dbf["RECCNUM"]?.ToString(), out var recnum) ? recnum : 0,
                    Locked = dbf["LOCKED"]?.ToString() == "TRUE",
                    Remark = dbf["REMARK"]?.ToString()
                };

                fastCvs.Add(fastCv);
            }

            return fastCvs;
        }


        public List<Company> GetCompaniesFromDbf()
        {
            var companies = new List<Company>();
            using var dbf = new DbfDataReader.DbfDataReader(DbfPaths.CompanyPath);

            while (dbf.Read())
            {
                var company = new Company
                {
                    Code = dbf["CODE"]?.ToString(),
                    BankCode = dbf["BANKCODE"]?.ToString(),
                    Co = dbf["CO"]?.ToString(),
                    BankName = dbf["BANKNAME"]?.ToString()
                };

                companies.Add(company);
            }

            return companies;
        }
    }
}
