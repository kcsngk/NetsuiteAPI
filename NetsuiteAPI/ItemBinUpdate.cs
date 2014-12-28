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
    public class ItemBinUpdate
    {
        public Item SKU { get; set; }
        public InventoryBin Bin { get; set; }
        public List<Item> ItemsToAdd { get; set; }

        public ItemBinUpdate()
        {
        }

        /*
        public List<Item> GetItemsFromExcel(Excel._Worksheet excelWorksheet)
        {
            List<Item> itemsToUpdate = new List<Item>();
            Excel.Range excelRange = excelWorksheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            int columnCount = excelRange.Columns.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                if (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null)
                {
                    string itemName = excelRange.Cells[i, 1].Value2.ToString();
                    itemsToUpdate.Add(new Item(itemName));
                }
            }

            return itemsToUpdate;

        }

        public List<InventoryBin> GetBinsFromExcel(Excel._Worksheet excelWorksheet)
        {
            List<InventoryBin> binsToAdd = new List<InventoryBin>();
            Excel.Range excelRange = excelWorksheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            int columnCount = excelRange.Columns.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                if (excelRange.Cells[i, 2] != null && excelRange.Cells[i, 2].Value2 != null && excelRange.Cells[i, 3] != null && excelRange.Cells[i, 3].Value2 != null)
                {
                    string location = excelRange.Cells[i, 2].Value2.ToString();
                    string binNumber = excelRange.Cells[i, 3].Value2.ToString();
                    binsToAdd.Add(new InventoryBin(location,binNumber));
                }
            }

            return binsToAdd;

        }
        */
        public List<InventoryBin> GetItemBinsFromExcel(string workbook)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbook);
            Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
            Excel.Range excelRange = excelWorksheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            int columnCount = excelRange.Columns.Count;

            List<InventoryBin> itemBins = new List<InventoryBin>();
            for (int i = 2; i <= rowCount; i++)
            {
                if (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null && excelRange.Cells[i, 2] != null && excelRange.Cells[i, 2].Value2 != null && excelRange.Cells[i, 3] != null && excelRange.Cells[i, 3].Value2 != null)
                {
                    string location = excelRange.Cells[i, 2].Value2.ToString();
                    string binNumber = excelRange.Cells[i, 3].Value2.ToString();
                    string itemName = excelRange.Cells[i, 1].Value2.ToString();
                    Item itemToAdd = new Item(itemName); 
                    itemBins.Add(new InventoryBin(location,binNumber,itemToAdd));
                }
            }

            return itemBins;

        }

        public string addBin(InventoryBin itemBin)
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            SearchStringField objItemName = new SearchStringField();
            objItemName.searchValue = itemBin.InventoryItem.ItemName;
            objItemName.@operator = SearchStringFieldOperator.@is;
            objItemName.operatorSpecified = true;
            ItemSearch objItemSearch = new ItemSearch();
            objItemSearch.basic = new ItemSearchBasic();
            objItemSearch.basic.itemId = objItemName;
            SearchResult objItemResult = service.search(objItemSearch);

            if (objItemResult.status.isSuccess != true)
            {
                Console.WriteLine("Cannot find Item " + itemBin.InventoryItem.ItemName + " " + objItemResult.status.statusDetail[0].message);
                return "Cannot find Item " + itemBin.InventoryItem.ItemName + " " + objItemResult.status.statusDetail[0].message;
            }
            if (objItemResult.recordList.Count() != 1)
            {
                Console.WriteLine("More than one item found for item " + itemBin.InventoryItem.ItemName);
                return "More than one item found for item " + itemBin.InventoryItem.ItemName;
            }


            SearchStringField objBinNumber = new SearchStringField();
            objBinNumber.searchValue = itemBin.BinNumber;
            objBinNumber.@operator = SearchStringFieldOperator.@is;
            objBinNumber.operatorSpecified = true;
            BinSearch objBinSearch = new BinSearch();
            objBinSearch.basic = new BinSearchBasic();
            objBinSearch.basic.binNumber = objBinNumber;
            SearchResult objBinResult = service.search(objBinSearch);

            if (objBinResult.status.isSuccess == false)
            {
                Console.WriteLine("Unable to find bin " + itemBin.BinNumber + " in NetSuite");
                return "Unable to find bin " + itemBin.BinNumber + " in NetSuite";
            }
            if (objItemResult.recordList.Count() != 1)
            {
                Console.WriteLine("More than one bin found for " + itemBin.BinNumber);
                return "More than one bin found for " + itemBin.BinNumber;
            }
            if (((Bin)objBinResult.recordList[0]).isInactive)
            {
                Console.WriteLine("Bin Number " + itemBin.BinNumber + " is inactive in NetSuite");
                return "Bin Number " + itemBin.BinNumber + " is inactive in NetSuite";
            }

            RecordRef objAddBin = new RecordRef();
            objAddBin.type = RecordType.bin;
            objAddBin.typeSpecified = true;
            objAddBin.internalId = ((Bin)objBinResult.recordList[0]).internalId;
            objAddBin.name = ((Bin)objBinResult.recordList[0]).binNumber;

            InventoryItemBinNumber objItemBinNumber = new InventoryItemBinNumber();
            objItemBinNumber.location = ((Bin)objBinResult.recordList[0]).location.internalId;
            objItemBinNumber.binNumber = objAddBin;

            InventoryItemBinNumber[] objItemBinNumbers = new InventoryItemBinNumber[1];
            objItemBinNumbers[0] = objItemBinNumber;

             if(((InventoryItem)objItemResult.recordList[0]).useBins==true)
                    {
                    
                    //Initialize item to update
                    InventoryItem objInventoryItem = new InventoryItem();
                    objInventoryItem.internalId = ((InventoryItem)objItemResult.recordList[0]).internalId;
                    objInventoryItem.salesDescription = ((InventoryItem)objItemResult.recordList[0]).salesDescription;
                    objInventoryItem.purchaseDescription = ((InventoryItem)objItemResult.recordList[0]).purchaseDescription;
                    
                    objInventoryItem.binNumberList = new InventoryItemBinNumberList();
                    objInventoryItem.binNumberList.binNumber = objItemBinNumbers;
                    objInventoryItem.binNumberList.replaceAll = false;
                    

                    //Request to update item with bins
                    
                    
                    WriteResponse objWriteResponse = service.update(objInventoryItem);

                    if (objWriteResponse.status.isSuccess != true)
                    {
                        Console.WriteLine(objWriteResponse.status.statusDetail[0].message);
                        return objWriteResponse.status.statusDetail[0].message;
                    }

                    Console.WriteLine("Bins successfully updated for: "+((InventoryItem)objItemResult.recordList[0]).itemId);
                    return "Bins successfully updated for: " + ((InventoryItem)objItemResult.recordList[0]).itemId;
                    }
             return string.Empty;
                    
        }
         
    }
}
