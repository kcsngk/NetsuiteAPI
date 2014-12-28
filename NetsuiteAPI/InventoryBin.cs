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
    public class InventoryBin
    {
        public string LocationName { get; set; }
        public string BinNumber { get; set; }
        public Item InventoryItem { get; set; }
        public int AdjustmentQuantity { get; set; }
        public Bin binRecord { get; set; }
      
        public InventoryBin()
        {
        }

        public InventoryBin(string locationName, string binNumber)
        {
            this.LocationName = locationName;
            this.BinNumber = binNumber;
        }

        public InventoryBin(string locationName, string binNumber, Item item)
        {
            this.LocationName = locationName;
            this.BinNumber = binNumber;
            this.InventoryItem = item;
        }

        public InventoryBin(int quantity, string binNumber, Item item)
        {
            this.AdjustmentQuantity = Convert.ToInt16(quantity);
            this.BinNumber = binNumber;
            this.InventoryItem = item;
        }

        public void addBinRecord(string binNumber)
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            SearchStringField objBinNumber = new SearchStringField();
            objBinNumber.searchValue = binNumber;
            objBinNumber.@operator = SearchStringFieldOperator.@is;
            objBinNumber.operatorSpecified = true;
            BinSearch objBinSearch = new BinSearch();
            objBinSearch.basic = new BinSearchBasic();
            objBinSearch.basic.binNumber = objBinNumber;
            SearchResult objBinResult = service.search(objBinSearch);

            if (objBinResult.status.isSuccess == false) throw new Exception("Unable to find bin " + binNumber + " is NetSuite");
            if (((Bin)objBinResult.recordList[0]).isInactive) throw new Exception("Bin Number " + binNumber + " is inactive in NetSuite");

            this.binRecord = new Bin();
            this.binRecord = (Bin)objBinResult.recordList[0];
        }
    
   


    }
}
