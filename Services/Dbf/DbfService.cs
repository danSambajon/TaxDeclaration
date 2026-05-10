using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.VisualBasic;
using Npgsql;
using System.Data;
using System.Data.OleDb;
using System.Text;
using TaxDeclaration.Models;
using TaxDeclaration.Models.ViewModels;
using TaxDeclaration.Utilities.Constants;
using static TaxDeclaration.Models.ViewModels.DCRMainDisbursementViewModel;

namespace TaxDeclaration.Services.Dbf
{
    public class DbfService
    {
        public List<FastStationViewModel> GetStationsFromDbf()
        {

            var stations = new List<FastStationViewModel>();

            if (!File.Exists(DbfPaths.StationPath))
            {
                throw new FileNotFoundException($"DBF not found: {DbfPaths.StationPath}");
            }

            using var dbf = new DbfDataReader.DbfDataReader(DbfPaths.StationPath);

            var recordsFound = dbf.DbfTable.Header.RecordCount;
            var fields = dbf.DbfTable.Columns.Count;

            if (dbf.DbfTable.Header.RecordCount == 0)
            {
                Console.WriteLine("DBF file has 0 records.");
                throw new IndexOutOfRangeException("DBF file has 0 records.");
            }

            while (dbf.Read())
            {
                var station = new FastStationViewModel
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

        public List<FastCvViewModel> GetFastCvsFromDbf()
        {
            var fastCvs = new List<FastCvViewModel>();

            if (!File.Exists(DbfPaths.FastCvPath))
            {
                throw new FileNotFoundException($"DBF not found: {DbfPaths.FastCvPath}");
            }

            using var dbf = new DbfDataReader.DbfDataReader(DbfPaths.FastCvPath);

            var recordsFound = dbf.DbfTable.Header.RecordCount;
            var fields = dbf.DbfTable.Columns.Count;

            while (dbf.Read())
            {
                var fastCv = new FastCvViewModel
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

        public List<TaxDeclareCompanyViewModel> GetCompaniesFromDbf()
        {
            var companies = new List<TaxDeclareCompanyViewModel>();

            if (!File.Exists(DbfPaths.CompanyPath))
            {
                throw new FileNotFoundException($"DBF not found: {DbfPaths.CompanyPath}");
            }

            using var dbf = new DbfDataReader.DbfDataReader(DbfPaths.CompanyPath);

            var recordsFound = dbf.DbfTable.Header.RecordCount;
            var fields = dbf.DbfTable.Columns.Count;

            while (dbf.Read())
            {
                var company = new TaxDeclareCompanyViewModel
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

        public List<DCRMainDisbursementViewModel> GetDCRMainDisbursementsFromDbf()
        {
            var records = new List<DCRMainDisbursementViewModel>();

            if (!File.Exists(DbfPaths.DcrMainDisbursementDbfPath))
                throw new FileNotFoundException($"DBF not found: {DbfPaths.DcrMainDisbursementDbfPath}");

            using var dbf = new DbfDataReader.DbfDataReader(DbfPaths.DcrMainDisbursementDbfPath);

            while (dbf.Read())
            {
                if (dbf["CASHPODATE"] != null || dbf["DCRDATE"] != null)
                {
                    var record = new DCRMainDisbursementViewModel
                    {
                        StnCode = dbf["STNCODE"]?.ToString(),
                        AccountNo = dbf["ACCOUNTNO"]?.ToString(),
                        CashPoDate = ToDate(dbf["CASHPODATE"]),
                        DcrDate = ToDate(dbf["DCRDATE"])
                    };

                    records.Add(record);
                }
            }

            return records;
        }

        public void GetBirSysFromDbf()
        {
            var twoThreeOSeven = GetTwoThreeOSeven();
            var atc = GetATC();
            var payee = GetPayee();
            var company = GetEwtCompany();

            //var cur2307 = (
            //    from a in twoThreeOSeven
            //    join b in atc on a.AtcCode equals b. into bGroup
            //    from b in bGroup.DefaultIfEmpty()
            //    join c in payee on a.payeecode equals c. into cGroup
            //    from c in cGroup.DefaultIfEmpty()
            //    join d in company on a.companycode.Trim() equals d.Code.Trim() into dGroup
            //    from d in dGroup.DefaultIfEmpty()
            //    where !string.IsNullOrWhiteSpace(a.rmonth)
            //       && !string.IsNullOrWhiteSpace(a.ryear)
            //       && !string.IsNullOrWhiteSpace(a.payeecode)
            //       && (a.downloadfrom != null && a.downloadfrom.Contains("IBS"))
            //    select new
            //    {
            //        a.datefrom,
            //        a.dateto,
            //        a.cvno,
            //        a.voucher_no,
            //        a.atc_code,
            //        a.total,
            //        a.amount,
            //        a.stationcode,
            //        a.approved,
            //        a.rmonth,
            //        a.ryear,
            //        desc = (b == null || b.desc == null) ? new string(' ', 254) : b.desc,
            //        percent = (b == null || b.percent == null) ? 0 : b.percent,
            //        a.downloadfrom,
            //        a.cancelled,
            //        a.payeecode,
            //        payeename = c == null ? null : c.payeename,
            //        tin_a = c == null ? null : c.tin_a,
            //        tin_b = c == null ? null : c.tin_b,
            //        tin_c = c == null ? null : c.tin_c,
            //        tin_d = c == null ? null : c.tin_d,
            //        payorcode = d == null ? null : d.code,
            //        payorname = d == null ? null : d.name,
            //        a.description
            //    })
            //    .Distinct()
            //    .OrderBy(x => x.ryear)
            //    .ThenBy(x => x.rmonth)
            //    .ThenBy(x => x.cvno)
            //    .ToList();



        }

        public List<TwoThreeOSevenViewModel> GetTwoThreeOSeven()
        {
            var twoThreeOSeven = new List<TwoThreeOSevenViewModel>();
            var problematicRows = new List<(int Row, string Error)>();

            if (!File.Exists(DbfPaths.Ewt2307TwoThreeOSevenDbfPath))
                throw new FileNotFoundException($"DBF not found: {DbfPaths.Ewt2307TwoThreeOSevenDbfPath}");

            // =============================================
            // PASS 1 — Normal read, collect problematic row numbers
            // =============================================
            using (var dbf1 = new DbfDataReader.DbfDataReader(DbfPaths.Ewt2307TwoThreeOSevenDbfPath))
            {
                int row = 0;
                while (true)
                {
                    row++;
                    bool hasRow;

                    try
                    {
                        hasRow = dbf1.Read();
                    }
                    catch (Exception ex)
                    {
                        problematicRows.Add((row, ex.Message));
                        continue;
                    }

                    if (!hasRow) break;

                    try
                    {
                        var record = new TwoThreeOSevenViewModel
                        {
                            RecId = int.Parse(dbf1["RECID"]?.ToString()),
                            ControlNum = dbf1["CONTROLNUM"]?.ToString(),
                            DateFrom = ToDate(dbf1["DATEFROM"]),
                            DateTo = ToDate(dbf1["DATETO"]),
                            VoucherNo = dbf1["VOUCHER_NO"]?.ToString(),
                            CvNo = dbf1["CVNO"]?.ToString(),
                            PayeeCode = dbf1["PAYEECODE"]?.ToString(),
                            AtcCode = dbf1["ATC_CODE"]?.ToString(),
                            Total = decimal.TryParse(dbf1["TOTAL"]?.ToString(), out var total) ? total : (decimal?)null,
                            Amount = decimal.TryParse(dbf1["AMOUNT"]?.ToString(), out var amount) ? amount : (decimal?)null,
                            DownloadFrom = dbf1["DOWNLOADFR"]?.ToString(),
                            StationCode = dbf1["STATIONCOD"]?.ToString(),
                            Description = dbf1["DESCRIPTIO"]?.ToString(),
                            Approved = bool.TryParse(dbf1["APPROVED"]?.ToString(), out var approved) ? approved : false,
                            RMonth = int.TryParse(dbf1["RMONTH"]?.ToString(), out var rmonth) ? rmonth : (int?)null,
                            RYear = int.TryParse(dbf1["RYEAR"]?.ToString(), out var ryear) ? ryear : (int?)null,
                            Cancelled = bool.TryParse(dbf1["CANCELLED"]?.ToString(), out var cancelled) ? cancelled : false
                        };

                        twoThreeOSeven.Add(record);
                    }
                    catch (Exception ex)
                    {
                        problematicRows.Add((row, ex.Message));
                    }
                }
            }

            // =============================================
            // PASS 2 — Read field layout then recover problematic rows via raw bytes
            // =============================================
            if (problematicRows.Any())
            {
                // build field map from DBF header
                var fields = new List<(string Name, char Type, int Size, int Offset)>();
                int headerLength;
                int recordLength;

                using (var headerStream = new FileStream(DbfPaths.Ewt2307TwoThreeOSevenDbfPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // read header length and record length from bytes 8-11
                    var mainHeader = new byte[32];
                    headerStream.Read(mainHeader, 0, 32);
                    headerLength = mainHeader[8] | (mainHeader[9] << 8);
                    recordLength = mainHeader[10] | (mainHeader[11] << 8);

                    // read field descriptors
                    headerStream.Seek(32, SeekOrigin.Begin);
                    int fieldOffset = 1; // after deletion flag
                    while (true)
                    {
                        var descriptor = new byte[32];
                        headerStream.Read(descriptor, 0, 32);
                        if (descriptor[0] == 0x0D) break;

                        var name = Encoding.ASCII.GetString(descriptor, 0, 11).TrimEnd('\0');
                        var type = (char)descriptor[11];
                        var size = descriptor[16];

                        fields.Add((name, type, size, fieldOffset));
                        fieldOffset += size;
                    }
                }

                // recover each problematic row
                using var rawStream = new FileStream(DbfPaths.Ewt2307TwoThreeOSevenDbfPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                foreach (var (rowNum, error) in problematicRows)
                {
                    try
                    {
                        long rowOffset = headerLength + ((long)(rowNum - 1) * recordLength);
                        rawStream.Seek(rowOffset, SeekOrigin.Begin);
                        var buffer = new byte[recordLength];
                        rawStream.Read(buffer, 0, buffer.Length);

                        var recovered = ParseRawRow(buffer, fields);
                        twoThreeOSeven.Add(recovered);
                    }
                    catch (Exception ex)
                    {
                        // truly unrecoverable
                        Console.WriteLine($"[UNRECOVERABLE] Row {rowNum} | {ex.Message}");
                    }
                }
            }

            return twoThreeOSeven;
        }

        public List<ATCViewModel> GetATC()
        {
            var records = new List<ATCViewModel>();

            if (!File.Exists(DbfPaths.Ewt2307AtcDbfPath))
                throw new FileNotFoundException($"DBF not found: {DbfPaths.Ewt2307AtcDbfPath}");

            using var dbf = new DbfDataReader.DbfDataReader(DbfPaths.Ewt2307AtcDbfPath);

            while (dbf.Read())
            {
                var record = new ATCViewModel
                {
                    Desc = dbf["DESC"]?.ToString(),
                    Percent = decimal.TryParse(dbf["PERCENT"]?.ToString(), out var percent) ? percent : (decimal?)null,
                };

                records.Add(record);
            }

            return records;
        }

        public List<PayeeViewModel> GetPayee()
        {
            var records = new List<PayeeViewModel>();

            if (!File.Exists(DbfPaths.Ewt2307PayeeDbfPath))
                throw new FileNotFoundException($"DBF not found: {DbfPaths.Ewt2307PayeeDbfPath}");

            using var dbf = new DbfDataReader.DbfDataReader(DbfPaths.Ewt2307PayeeDbfPath);

            while (dbf.Read())
            {
                var record = new PayeeViewModel
                {
                    PayeeName = dbf["PAYEENAME"]?.ToString(),
                    PayeeCode = dbf["PAYEECODE"]?.ToString(),
                    TinA = dbf["TIN_A"]?.ToString(),
                    TinB = dbf["TIN_B"]?.ToString(),
                    TinC = dbf["TIN_C"]?.ToString(),
                    TinD = dbf["TIN_D"]?.ToString(),
                };

                records.Add(record);
            }

            return records;
        }

        public List<EwtCompanyViewModel> GetEwtCompany()
        {
            var records = new List<EwtCompanyViewModel>();

            if (!File.Exists(DbfPaths.Ewt2307EwtCompanyDbfPath))
                throw new FileNotFoundException($"DBF not found: {DbfPaths.Ewt2307EwtCompanyDbfPath}");

            using var dbf = new DbfDataReader.DbfDataReader(DbfPaths.Ewt2307EwtCompanyDbfPath);

            while (dbf.Read())
            {
                var record = new EwtCompanyViewModel
                {
                    Code = dbf["CODE"]?.ToString(),
                    Name = dbf["NAME"]?.ToString(),
                };

                records.Add(record);
            }

            return records;
        }

        #region == Helpers ==

        private static DateTime? ToDate(object value) =>
            value != null && value != DBNull.Value && DateTime.TryParse(value.ToString(), out var d)
                ? d : null;

        TwoThreeOSevenViewModel ParseRawRow(byte[] rowBytes, List<(string Name, char Type, int Size, int Offset)> fields)
        {
            var record = new TwoThreeOSevenViewModel();

            var debugLog = new StringBuilder();
            debugLog.AppendLine($"Buffer length: {rowBytes.Length}");
            debugLog.AppendLine("Field offsets:");
            foreach (var (name, type, size, offset) in fields)
            {
                debugLog.AppendLine($"  {name} | offset: {offset} | size: {size} | end: {offset + size}");
            }
            var debugOutput = debugLog.ToString();
            // breakpoint here

            foreach (var (name, type, size, offset) in fields)
            {
                try
                {
                    var fieldBytes = rowBytes.AsSpan(offset, size);
                    var raw = Encoding.ASCII.GetString(fieldBytes.ToArray()).Trim();

                    switch (name)
                    {
                        // D type — the crashers, try to parse, null if blank
                        case "DATEFROM": record.DateFrom = ParseDbfDate(raw); break;
                        case "DATETO": record.DateTo = ParseDbfDate(raw); break;

                        // everything else
                        case "RECID": record.RecId = BitConverter.ToInt32(fieldBytes.ToArray(), 0); break;
                        case "CONTROLNUM": record.ControlNum = raw; break;
                        case "VOUCHER_NO": record.VoucherNo = raw; break;
                        case "CVNO": record.CvNo = raw; break;
                        case "PAYEECODE": record.PayeeCode = raw; break;
                        case "ATC_CODE": record.AtcCode = raw; break;
                        case "TOTAL": record.Total = string.IsNullOrEmpty(raw) ? 0 : decimal.Parse(raw); break;
                        case "AMOUNT": record.Amount = string.IsNullOrEmpty(raw) ? 0 : decimal.Parse(raw); break;
                        case "DOWNLOADFR": record.DownloadFrom = raw; break;
                        case "STATIONCOD": record.StationCode = raw; break;
                        case "DESCRIPTIO": record.Description = raw; break;
                        case "APPROVED": record.Approved = raw == "T"; break;
                        case "RMONTH": record.RMonth = string.IsNullOrEmpty(raw) ? 0 : int.Parse(raw); break;
                        case "RYEAR": record.RYear = string.IsNullOrEmpty(raw) ? 0 : int.Parse(raw); break;
                        case "CANCELLED": record.Cancelled = raw == "T"; break;
                    }
                }
                catch
                {
                    // this individual field failed — leave as default/null, move to next field
                }
            }

            return record;
        }

        DateTime? ParseDbfDate(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw) || raw.Trim('0').Length == 0)
                return null;
            if (DateTime.TryParseExact(raw, "yyyyMMdd", null,
                System.Globalization.DateTimeStyles.None, out var result))
                return result;
            return null;
        }

        #endregion == Helpers ==
    }
}
