using NetsuiteAPI.com.netsuite.webservices;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace NetsuiteAPI
{
    public class ItemFulfillment
    {
        public com.netsuite.webservices.ItemFulfillment itemFulfillment { get; set; }

        public ItemFulfillment()
        {
        }

        public ItemFulfillment(SalesOrder salesOrder)
        {
            itemFulfillment = new com.netsuite.webservices.ItemFulfillment();
            RecordRef salesOrderID = new RecordRef();
            salesOrderID.type = RecordType.salesOrder;
            salesOrderID.typeSpecified = true;
            salesOrderID.internalId = salesOrder.salesOrder.internalId;

            itemFulfillment.createdFrom = salesOrderID;

            ItemFulfillmentItemList itemFulItemList = new ItemFulfillmentItemList();
            int i = 0;
            ItemFulfillmentItem[] Items = new ItemFulfillmentItem[salesOrder.salesOrder.itemList.item.Count()];

            Dictionary<string, double> binQuantities = new Dictionary<string, double>();

            foreach (var itemLine in salesOrder.salesOrder.itemList.item)
            {
                Items[i] = new ItemFulfillmentItem();
                RecordRef item = new RecordRef();
                item.type = RecordType.inventoryItem;
                item.typeSpecified = true;
                item.internalId = itemLine.item.internalId;
                Items[i].item = item;
                Items[i].quantity = itemLine.quantity;
                Items[i].quantitySpecified = true;
                Items[i].orderLine = itemLine.line;
                Items[i].orderLineSpecified = true;

                binQuantities = getBins(itemLine.item.name);

                string[] bins = new string[binQuantities.Count];
                int j = 0;
                foreach (var key in binQuantities)
                {
                    bins[j] = key.Key;
                    j++;
                }

                Array.Sort(bins, new AlphanumComparatorFast());
                bins = bins.Where(x => !string.IsNullOrEmpty(x)).ToArray();

                List<string> notPreferredBins = new List<string>();
                List<string> preferredBins = new List<string>();
                List<string> orderedBins = new List<string>();

                Regex regex = new Regex("[0-9]+[A-Z]");
                string C = "C";

                foreach (var bin in bins)
                {
                    Match match = regex.Match(bin);
                    if (match.Success)
                    {
                        string letter = (match.Value[match.Value.Length - 1]).ToString();
                        int compare = string.Compare(letter, C);
                        if (compare < 0)
                        {
                            notPreferredBins.Add(bin);
                        }
                        else preferredBins.Add(bin);
                    }
                }

                foreach (var bin in preferredBins)
                {
                    orderedBins.Add(bin);
                }
                orderedBins.Add("Stock");
                foreach (var bin in notPreferredBins)
                {
                    orderedBins.Add(bin);
                }

                foreach (var bin in orderedBins)
                {
                    double availQuantity = binQuantities[bin];
                    if (availQuantity >= Items[i].quantity)
                    {
                        Items[i].binNumbers = bin;
                        break;
                    }
                }
                i++;
            }
            itemFulItemList.item = Items;
            this.itemFulfillment.itemList = itemFulItemList;
        }

        public Dictionary<string, double> getBins(string SKU)
        {
            NetSuiteService objService = new NetSuiteService();
            objService.CookieContainer = new CookieContainer();
            Passport passport = new Passport();
            passport.account = "3451682";
            passport.email = "kevin.ng@tridentcase.com";
            RecordRef role = new RecordRef();
            role.internalId = "1026";
            passport.role = role;
            passport.password = "tridenT168";
            Passport objPassport = passport;
            Status objStatus = objService.login(objPassport).status;

            ItemSearchAdvanced isa = new ItemSearchAdvanced();
            isa.savedSearchId = "141";  //substitute your own saved search internal ID
            ItemSearch iS = new ItemSearch();
            ItemSearchBasic isb = new ItemSearchBasic();

            SearchStringField itemSKU = new SearchStringField();
            itemSKU.searchValue = SKU;
            itemSKU.@operator = SearchStringFieldOperator.contains;
            itemSKU.operatorSpecified = true;
            isb.itemId = itemSKU;
            iS.basic = isb;
            isa.criteria = iS;

            SearchResult sr = new SearchResult();
            sr = objService.search(isa);

            if (sr.status.isSuccess != true) throw new Exception("Cannot find item.");

            Dictionary<string, double> binNumberList = new Dictionary<string, double>();

            foreach (ItemSearchRow irow in sr.searchRowList)
            {
                if (irow.basic.itemId[0].searchValue == SKU)
                {
                    binNumberList.Add(irow.basic.binNumber[0].searchValue, irow.basic.binOnHandAvail[0].searchValue);
                }
            }

            return binNumberList;
        }

        public List<ItemFulfillmentLine> createNewItemFulfillmentLines(ItemFulfillmentItemList oldItemList)
        {
            List<ItemFulfillmentLine> newItemList = new List<ItemFulfillmentLine>();
            bool flag = false;
            foreach (var line in oldItemList.item)
            {
                ItemFulfillmentLine newLine = new ItemFulfillmentLine(line.item.name, line.quantity, flag);
                newItemList.Add(newLine);
            }

            return newItemList;
        }

        public void itemFulfillmentSearch(string orderNumber)
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            TransactionSearchAdvanced tranSearchAdvanced = new TransactionSearchAdvanced();
            TransactionSearch tranSearch = new TransactionSearch();
            SearchStringField orderID = new SearchStringField();
            orderID.@operator = SearchStringFieldOperator.@is;
            orderID.operatorSpecified = true;
            orderID.searchValue = orderNumber;

            TransactionSearchBasic tranSearchBasic = new TransactionSearchBasic();
            tranSearchBasic.tranId = orderID;
            tranSearch.basic = tranSearchBasic;
            tranSearchAdvanced.criteria = tranSearch;

            SearchPreferences searchPreferences = new SearchPreferences();
            searchPreferences.bodyFieldsOnly = false;
            service.searchPreferences = searchPreferences;

            SearchResult transactionResult = service.search(tranSearchAdvanced);

            if (transactionResult.status.isSuccess != true) throw new Exception("Cannot find Order " + orderNumber + " " + transactionResult.status.statusDetail[0].message);
            if (transactionResult.recordList.Count() != 1) throw new Exception("More than one order found for item " + orderNumber);

            this.itemFulfillment = new com.netsuite.webservices.ItemFulfillment();
            this.itemFulfillment = ((com.netsuite.webservices.ItemFulfillment)transactionResult.recordList[0]);
        }

        public string markAsPacked()
        {
            ItemFulfillmentShipStatus packed = ItemFulfillmentShipStatus._packed;

            ItemFulfillment updatedOrder = new ItemFulfillment();
            updatedOrder.itemFulfillment = new com.netsuite.webservices.ItemFulfillment();
            updatedOrder.itemFulfillment.internalId = this.itemFulfillment.internalId;
            updatedOrder.itemFulfillment.shipStatus = packed;
            updatedOrder.itemFulfillment.shipStatusSpecified = true;

            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            WriteResponse writeResponse = service.update(updatedOrder.itemFulfillment);
            if (writeResponse.status.isSuccess == true)
            {
                Console.WriteLine("Item Fulfillment " + this.itemFulfillment.tranId + " marked as packed.");
            }
            if (writeResponse.status.isSuccess == false)
            {
                Console.WriteLine(writeResponse.status.statusDetail[0].message);
            }

            return this.itemFulfillment.tranId;
        }

        public string markAsShipped()
        {
            ItemFulfillmentShipStatus shipped = ItemFulfillmentShipStatus._shipped;

            ItemFulfillment updatedOrder = new ItemFulfillment();
            updatedOrder.itemFulfillment = new com.netsuite.webservices.ItemFulfillment();
            updatedOrder.itemFulfillment.internalId = this.itemFulfillment.internalId;
            updatedOrder.itemFulfillment.shipStatus = shipped;
            updatedOrder.itemFulfillment.shipStatusSpecified = true;

            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            WriteResponse writeResponse = service.update(updatedOrder.itemFulfillment);
            if (writeResponse.status.isSuccess == true)
            {
                Console.WriteLine("Item Fulfillment " + this.itemFulfillment.tranId + " marked as shipped.");
            }
            if (writeResponse.status.isSuccess == false)
            {
                Console.WriteLine(writeResponse.status.statusDetail[0].message);
            }

            return this.itemFulfillment.tranId;
        }

        public ItemFulfillmentItemList addToExistingItemList(ItemFulfillmentItemList oldList, ItemFulfillmentLine fulfillmentLine)
        {
            int oldCount = oldList.item.Length;
            int newCount = oldCount + 1;
            ItemFulfillmentItemList newFulfillmentList = new ItemFulfillmentItemList();
            ItemFulfillmentItem[] newItem = new ItemFulfillmentItem[1];
            ItemFulfillmentItem[] newList = new ItemFulfillmentItem[newCount];
            oldList.item.CopyTo(newList, 0);

            newItem[0] = new ItemFulfillmentItem();
            RecordRef item = new RecordRef();
            item.type = RecordType.inventoryItem;
            item.typeSpecified = true;
            item.name = fulfillmentLine.item.ItemName;
            item.internalId = fulfillmentLine.item.itemRecord.internalId;
            newItem[0].item = item;

            newItem[0].quantity = fulfillmentLine.QuantityRequested;
            newItem[0].quantitySpecified = true;

            newItem.CopyTo(newList, oldCount);
            newFulfillmentList.item = newList;

            return newFulfillmentList;
        }

        public void editFulfillment(List<ItemFulfillmentLine> fulfillmentLines)
        {
            ItemFulfillment updatedFulfillment = new ItemFulfillment();
            updatedFulfillment.itemFulfillment = new com.netsuite.webservices.ItemFulfillment();
            updatedFulfillment.itemFulfillment.internalId = this.itemFulfillment.internalId;
            updatedFulfillment.itemFulfillment.itemList = new ItemFulfillmentItemList();
            // updatedFulfillment.itemFulfillment.itemList.item = new ItemFulfillmentItem[this.itemFulfillment.itemList.item.Length];

            //bool flag = false;
            //bool contains = false;
            //bool contains1 = false;

            foreach (var line in fulfillmentLines)
            {
                foreach (var itemObject in this.itemFulfillment.itemList.item)
                {
                    //if (itemObject.item.name.Contains(line.item.ItemName))
                    if (itemObject.item.name.Contains(line.item.ItemName) && (line.delete == true))
                    {
                        this.itemFulfillment.itemList.item = this.itemFulfillment.itemList.item.Where(val => !val.item.name.Contains(line.item.ItemName)).ToArray();
                        //ItemFulfillmentItem ifi = new ItemFulfillmentItem();
                        //ifi.orderLine = itemObject.orderLine;
                    }
                    if (itemObject.item.name.Contains(line.item.ItemName) && (line.delete == false))
                    {
                        itemObject.quantity = line.QuantityRequested;
                    }
                    if (!itemObject.item.name.Contains(line.item.ItemName) && (line.delete == false))
                    {
                    }
                    if (!itemObject.item.name.Contains(line.item.ItemName) && (line.delete == true))
                    {
                    }
                }
            }

            updatedFulfillment.itemFulfillment.itemList = this.itemFulfillment.itemList;
            //updatedFulfillment.itemFulfillment.itemList.item = this.itemFulfillment.itemList.item;

            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            WriteResponse writeResponse = service.update(updatedFulfillment.itemFulfillment);
            if (writeResponse.status.isSuccess == true)
            {
                Console.WriteLine("Update Item Fulfillment success");
            }
            if (writeResponse.status.isSuccess == false)
            {
                Console.WriteLine(writeResponse.status.statusDetail[0].message);
            }
        }

        public string createItemFulfillment()
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            WriteResponse writeResponse = service.add(this.itemFulfillment);
            if (writeResponse.status.isSuccess == true)
            {
                Console.WriteLine("Create Item Fulfillment success");
                return "Create Item Fulfillment Success";
            }
            if (writeResponse.status.isSuccess == false)
            {
                Console.WriteLine(writeResponse.status.statusDetail[0].message);
                return writeResponse.status.statusDetail[0].message;
            }
            return string.Empty;
        }

        public string deleteFulfillment()
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            RecordRef orderToBeDeleted = new RecordRef();
            orderToBeDeleted.type = RecordType.itemFulfillment;
            orderToBeDeleted.typeSpecified = true;
            orderToBeDeleted.internalId = this.itemFulfillment.internalId;

            WriteResponse writeResponse = service.delete(orderToBeDeleted);
            if (writeResponse.status.isSuccess == true)
            {
                Console.WriteLine("Delete Item Fulfillment " + this.itemFulfillment.tranId + " success");
                return "Delete Item Fulfillment " + this.itemFulfillment.tranId + " success";
            }
            if (writeResponse.status.isSuccess == false)
            {
                Console.WriteLine(writeResponse.status.statusDetail[0].message);
                return writeResponse.status.statusDetail[0].message;
            }
            return string.Empty;
        }

        public string[] GetBinsFromExcel(string workbook)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbook);
            Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
            Excel.Range excelRange = excelWorksheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            int columnCount = excelRange.Columns.Count;
            string[] ordersToFulfill = new string[rowCount - 1];

            for (int i = 2; i <= rowCount; i++)
            {
                if (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null)
                {
                    string orderNumber = excelRange.Cells[i, 1].Value2.ToString();
                    ordersToFulfill[i - 2] = orderNumber;
                }
            }

            excelWorkbook.Close();
            return ordersToFulfill;
        }
    }

    public class AlphanumComparatorFast : IComparer
    {
        public int Compare(object x, object y)
        {
            string s1 = x as string;
            if (s1 == null)
            {
                return 0;
            }
            string s2 = y as string;
            if (s2 == null)
            {
                return 0;
            }

            int len1 = s1.Length;
            int len2 = s2.Length;
            int marker1 = 0;
            int marker2 = 0;

            // Walk through two the strings with two markers.
            while (marker1 < len1 && marker2 < len2)
            {
                char ch1 = s1[marker1];
                char ch2 = s2[marker2];

                // Some buffers we can build up characters in for each chunk.
                char[] space1 = new char[len1];
                int loc1 = 0;
                char[] space2 = new char[len2];
                int loc2 = 0;

                // Walk through all following characters that are digits or
                // characters in BOTH strings starting at the appropriate marker.
                // Collect char arrays.
                do
                {
                    space1[loc1++] = ch1;
                    marker1++;

                    if (marker1 < len1)
                    {
                        ch1 = s1[marker1];
                    }
                    else
                    {
                        break;
                    }
                } while (char.IsDigit(ch1) == char.IsDigit(space1[0]));

                do
                {
                    space2[loc2++] = ch2;
                    marker2++;

                    if (marker2 < len2)
                    {
                        ch2 = s2[marker2];
                    }
                    else
                    {
                        break;
                    }
                } while (char.IsDigit(ch2) == char.IsDigit(space2[0]));

                // If we have collected numbers, compare them numerically.
                // Otherwise, if we have strings, compare them alphabetically.
                string str1 = new string(space1);
                string str2 = new string(space2);

                int result;

                if (char.IsDigit(space1[0]) && char.IsDigit(space2[0]))
                {
                    int thisNumericChunk = int.Parse(str1);
                    int thatNumericChunk = int.Parse(str2);
                    result = thisNumericChunk.CompareTo(thatNumericChunk);
                }
                else
                {
                    result = str1.CompareTo(str2);
                }

                if (result != 0)
                {
                    return result;
                }
            }
            return len1 - len2;
        }
    }
}