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
    public class Item
    {
        public string ItemName { get; set; }
        public InventoryItem itemRecord { get; set; }
        public Item(string ItemName)
        {
            this.ItemName = ItemName;
        }

        public Item(InventoryItem inventoryItem)
        {
            this.ItemName = inventoryItem.itemId;
            this.itemRecord = inventoryItem;
        }

        public void addItemRecord()
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            string itemName = this.ItemName;

            SearchStringField objItemName = new SearchStringField();
            objItemName.searchValue = itemName;
            objItemName.@operator = SearchStringFieldOperator.@is;
            objItemName.operatorSpecified = true;
            ItemSearch objItemSearch = new ItemSearch();
            objItemSearch.basic = new ItemSearchBasic();
            objItemSearch.basic.itemId = objItemName;

            SearchPreferences searchPreferences = new SearchPreferences();
            searchPreferences.bodyFieldsOnly = false;
            service.searchPreferences = searchPreferences;

            SearchResult objItemResult = service.search(objItemSearch);

            if (objItemResult.status.isSuccess != true) throw new Exception("Cannot find Item " + itemName + " " + objItemResult.status.statusDetail[0].message);
            if (objItemResult.recordList.Count() != 1) throw new Exception("More than one item found for item " + itemName);

            InventoryItem iRecord = new InventoryItem();

            iRecord = ((InventoryItem)objItemResult.recordList[0]);

            this.itemRecord = new InventoryItem();
            this.itemRecord=iRecord;
        }

    }
}
