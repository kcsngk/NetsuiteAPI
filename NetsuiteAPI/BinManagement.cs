using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using NetsuiteAPI.com.netsuite.webservices;
using System.Net;

namespace NetsuiteAPI
{
    public class BinManagement
    {
        public InventoryBin Bin { get; set; }

        public BinManagement()
        {
        }

       
        public string createNewBin(InventoryBin bin)
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;
 
            LocationSearch locationSearch = new LocationSearch();
            LocationSearchBasic locationSearchBasic = new LocationSearchBasic();
            SearchStringField locationName = new SearchStringField();
            locationName.searchValue = bin.LocationName;
            locationName.@operator = SearchStringFieldOperator.@is;
            locationName.operatorSpecified = true;
            locationSearchBasic.name = locationName;
            locationSearch.basic = locationSearchBasic;

            SearchResult sr = new SearchResult();
            sr = service.search(locationSearch);
            if (sr.status.isSuccess != true) Console.WriteLine(sr.status.statusDetail[0].message);

            Bin newBin = new Bin();
            RecordRef newLocation = new RecordRef();
            newLocation.type = RecordType.location;
            newLocation.typeSpecified = true;
            newLocation.internalId = ((com.netsuite.webservices.Location)sr.recordList[0]).internalId;
            newBin.binNumber = bin.BinNumber;
            newBin.location = newLocation;

            WriteResponse writeResponse = service.add(newBin);
            if (writeResponse.status.isSuccess == true)
            {

                Console.WriteLine("Bin: "+ newBin.binNumber +" has been created.");
                return "Bin: " + newBin.binNumber + " has been created.";

            }
            if (writeResponse.status.isSuccess == false)
            {
                Console.WriteLine(writeResponse.status.statusDetail[0].message);
                return writeResponse.status.statusDetail[0].message;
            }
            return string.Empty;
        }


        public void DeleteBin(InventoryBin bin)
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;


            BinSearch binSearch = new BinSearch();
            BinSearchBasic bSBasic = new BinSearchBasic();
            SearchStringField binName = new SearchStringField();
            binName.searchValue = bin.BinNumber;
            binName.@operator = SearchStringFieldOperator.@is;
            binName.operatorSpecified = true;
            bSBasic.binNumber = binName;
            binSearch.basic = bSBasic;

            SearchResult sr = new SearchResult();
            sr = service.search(binSearch);
            if (sr.status.isSuccess != true) Console.WriteLine(sr.status.statusDetail[0].message);

            RecordRef binToDelete = new RecordRef();
            binToDelete.type = RecordType.bin;
            binToDelete.typeSpecified = true;
            binToDelete.name = bin.BinNumber;
            binToDelete.internalId = ((Bin)sr.recordList[0]).internalId;


            WriteResponse writeResponse = service.delete(binToDelete);
            if (writeResponse.status.isSuccess == true)
            {
                Console.WriteLine("Bin: " + binToDelete.name + " has been deleted.");
            }
            if (writeResponse.status.isSuccess == false)
            {
                Console.WriteLine(writeResponse.status.statusDetail[0].message);

            }
        }


        public List<InventoryBin> GetBinsFromExcel(string workbook)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbook);
            Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
            Excel.Range excelRange = excelWorksheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            int columnCount = excelRange.Columns.Count;
            List<InventoryBin> binsToAdd = new List<InventoryBin>();

            for (int i = 2; i <= rowCount; i++)
            {
                if (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null && excelRange.Cells[i, 2] != null && excelRange.Cells[i, 2].Value2 != null)
                {
                    string location = excelRange.Cells[i, 1].Value2.ToString();
                    string binNumber = excelRange.Cells[i, 2].Value2.ToString();

                    binsToAdd.Add(new InventoryBin(location, binNumber));
                }
            }

            excelWorkbook.Close();
            return binsToAdd;

        }

     
    }
}
