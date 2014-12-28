using NetsuiteAPI.com.netsuite.webservices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace NetsuiteAPI
{
    public class SalesOrderInfo
    {
        public string shippingAddress { get; set; }

        public string billingAddress { get; set; }

        public string orderMemo { get; set; }

        public string fob { get; set; }

        public SalesOrderInfo(string shippingAddress, string billingAddress, string orderMemo, string fob)
        {
            this.shippingAddress = shippingAddress;
            this.billingAddress = billingAddress;
            this.orderMemo = orderMemo;
            this.fob = fob;
        }
    }

    public class SalesOrder
    {
        public string Customer { get; set; }

        public com.netsuite.webservices.SalesOrder salesOrder { get; set; }

        public string fob { get; set; }

        public SalesOrder()
        {
        }

        public SalesOrder(SalesOrderInfo salesOrderInfo)
        {
            this.salesOrder = new com.netsuite.webservices.SalesOrder();
            this.salesOrder.shipAddress = salesOrderInfo.shippingAddress;
            this.salesOrder.billAddress = salesOrderInfo.billingAddress;
            this.salesOrder.memo = salesOrderInfo.orderMemo;
            this.fob = salesOrderInfo.fob;
            // this.salesOrder.location = location.locationRecord;
        }

        public void addItemList(List<SalesOrderLine> salesOrderLines)
        {
            SalesOrderItemList salesOrderItemList = new SalesOrderItemList();

            int i = 0;
            SalesOrderItem[] Items = new SalesOrderItem[salesOrderLines.Count];
            foreach (var SOLine in salesOrderLines)
            {
                Items[i] = new SalesOrderItem();
                RecordRef item = new RecordRef();
                item.type = RecordType.inventoryItem;
                item.typeSpecified = true;
                item.internalId = salesOrderLines[i].item.itemRecord.internalId;
                Items[i].item = item;

                RecordRef prLevel = new RecordRef();
                prLevel.type = RecordType.priceLevel;
                prLevel.internalId = "-1";
                prLevel.typeSpecified = true;

                Items[i].price = prLevel;
                Items[i].rate = Convert.ToString(salesOrderLines[i].UnitPrice);
                Items[i].quantity = salesOrderLines[i].QuantityRequested;
                Items[i].quantitySpecified = true;
                i++;
            }
            salesOrderItemList.item = Items;
            this.salesOrder.itemList = salesOrderItemList;
        }

        public SalesOrderItemList addToExistingItemList(SalesOrderItemList oldList, SalesOrderLine salesOrderLine)
        {
            int oldCount = oldList.item.Length;
            int newCount = oldCount + 1;
            SalesOrderItemList newOrderList = new SalesOrderItemList();
            SalesOrderItem[] newItem = new SalesOrderItem[1];
            SalesOrderItem[] newList = new SalesOrderItem[newCount];
            oldList.item.CopyTo(newList, 0);

            newItem[0] = new SalesOrderItem();
            RecordRef item = new RecordRef();
            item.type = RecordType.inventoryItem;
            item.typeSpecified = true;
            // item.internalId = "1229";
            item.name = salesOrderLine.item.ItemName;
            item.internalId = salesOrderLine.item.itemRecord.internalId;
            newItem[0].item = item;

            RecordRef prLevel = new RecordRef();
            prLevel.type = RecordType.priceLevel;
            prLevel.internalId = "-1";
            prLevel.typeSpecified = true;

            newItem[0].price = prLevel;
            newItem[0].rate = Convert.ToString(salesOrderLine.UnitPrice);
            newItem[0].quantity = salesOrderLine.QuantityRequested;
            newItem[0].quantitySpecified = true;

            newItem.CopyTo(newList, oldCount);
            newOrderList.item = newList;

            return newOrderList;
        }

        public void setCustomer(string customerID)
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            CustomerSearch customerSearch = new CustomerSearch();
            customerSearch = CustomerNameSearch(customerID);
            SearchResult searchResult = service.search(customerSearch);
            if (searchResult.status.isSuccess)
            {
                if (searchResult.recordList != null && searchResult.recordList.Length == 1)
                {
                    string entityID = ((Customer)searchResult.recordList[0]).entityId;
                    Console.WriteLine(entityID);
                }
            }
            else
            {
                throw new Exception("Cannot find Customer " + customerID + " " + searchResult.status.statusDetail[0].message);
            }

            this.salesOrder.entity = createCustomer(searchResult);
        }

        public CustomerSearch CustomerNameSearch(string custName)
        {
            CustomerSearch custSearch = new CustomerSearch();
            SearchStringField customerEntityID = new SearchStringField();
            customerEntityID.@operator = SearchStringFieldOperator.@is;
            customerEntityID.operatorSpecified = true;
            customerEntityID.searchValue = custName;
            CustomerSearchBasic custBasic = new CustomerSearchBasic();
            custBasic.entityId = customerEntityID;
            custSearch.basic = custBasic;
            return custSearch;
        }

        public RecordRef createCustomer(SearchResult searchResult)
        {
            RecordRef customer = new RecordRef();
            customer.type = RecordType.customer;
            customer.typeSpecified = true;
            string entityID = ((Customer)searchResult.recordList[0]).entityId;
            customer.name = entityID;
            customer.internalId = ((Customer)searchResult.recordList[0]).internalId;
            return customer;
        }

        public List<SalesOrderLine> createNewSalesOrderLines(SalesOrderItemList oldItemList)
        {
            List<SalesOrderLine> newItemList = new List<SalesOrderLine>();
            bool flag = false;
            foreach (var line in oldItemList.item)
            {
                SalesOrderLine newLine = new SalesOrderLine(line.item.name, line.quantity, Convert.ToDecimal(line.rate), flag);
                newItemList.Add(newLine);
            }

            return newItemList;
        }

        public void salesOrderSearch(string orderNumber)
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

            this.salesOrder = new com.netsuite.webservices.SalesOrder();
            this.salesOrder = ((com.netsuite.webservices.SalesOrder)transactionResult.recordList[0]);
        }

        public bool placeSalesOrder()
        {
            bool success = false;
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            /*
            SelectCustomFieldRef tempCustFieldRef = new SelectCustomFieldRef();
            tempCustFieldRef.internalId = "148";
            tempCustFieldRef.scriptId = "custbodycustbodycust_fob";
            tempCustFieldRef.value.internalId = "3";
            this.salesOrder.customFieldList = new CustomFieldRef[4];
            this.salesOrder.customFieldList[4]  = new SelectCustomFieldRef();
            this.salesOrder.customFieldList[4] = tempCustFieldRef;
           */

            WriteResponse writeResponse = service.add(this.salesOrder);
            if (writeResponse.status.isSuccess == true)
            {
                success = true;
                Console.WriteLine("Sales Order success");
                Console.WriteLine(((RecordRef)writeResponse.baseRef).internalId);
                return success;
            }
            else
            {
                success = false;
                Console.WriteLine(writeResponse.status.statusDetail[0].message);
                return success;
            }
        }

        public void editOrder(List<SalesOrderLine> orderLines)
        {
            SalesOrder updatedOrder = new SalesOrder();
            updatedOrder.salesOrder = new com.netsuite.webservices.SalesOrder();
            updatedOrder.salesOrder.internalId = this.salesOrder.internalId;
            updatedOrder.salesOrder.itemList = new SalesOrderItemList();
            bool flag = false;
            bool contains = false;
            bool contains1 = false;

            foreach (var line in orderLines)
            {
                foreach (var itemObject in this.salesOrder.itemList.item)
                {
                    //if (itemObject.item.name.Contains(line.item.ItemName))
                    if (itemObject.item.name.Contains(line.item.ItemName) && (line.delete == true))
                    {
                        this.salesOrder.itemList.item = this.salesOrder.itemList.item.Where(val => !val.item.name.Contains(line.item.ItemName)).ToArray();
                    }
                    if (itemObject.item.name.Contains(line.item.ItemName) && (line.delete == false))
                    {
                        itemObject.quantity = line.QuantityRequested;
                        contains1 = true;
                    }
                    if (!itemObject.item.name.Contains(line.item.ItemName) && (line.delete == false))
                    {
                        contains = false;
                    }
                    if (!itemObject.item.name.Contains(line.item.ItemName) && (line.delete == true))
                    {
                        contains = true;
                    }
                }

                if (!(contains || contains1))
                {
                    SalesOrderLine addLine = new SalesOrderLine(line.item.ItemName, line.QuantityRequested, line.UnitPrice, flag);
                    this.salesOrder.itemList = this.addToExistingItemList(this.salesOrder.itemList, addLine);
                }
                contains = false;
                contains1 = false;
            }

            List<SalesOrderLine> newItemList = new List<SalesOrderLine>();
            newItemList = createNewSalesOrderLines(this.salesOrder.itemList);
            updatedOrder.addItemList(newItemList);

            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            WriteResponse writeResponse = service.update(updatedOrder.salesOrder);
            if (writeResponse.status.isSuccess == true)
            {
                Console.WriteLine("Update Sales Order success");
            }
            if (writeResponse.status.isSuccess == false)
            {
                Console.WriteLine(writeResponse.status.statusDetail[0].message);
            }
        }

        public void deleteOrder()
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            RecordRef orderToBeDeleted = new RecordRef();
            orderToBeDeleted.type = RecordType.salesOrder;
            orderToBeDeleted.typeSpecified = true;
            orderToBeDeleted.internalId = this.salesOrder.internalId;

            WriteResponse writeResponse = service.delete(orderToBeDeleted);
            if (writeResponse.status.isSuccess == true)
            {
                Console.WriteLine("Delete Sales Order " + this.salesOrder.tranId + " success");
            }
            if (writeResponse.status.isSuccess == false)
            {
                Console.WriteLine(writeResponse.status.statusDetail[0].message);
            }
        }

        public string approveOrder()
        {
            SalesOrder updatedOrder = new SalesOrder();
            updatedOrder.salesOrder = new com.netsuite.webservices.SalesOrder();
            updatedOrder.salesOrder.internalId = this.salesOrder.internalId;
            updatedOrder.salesOrder.itemList = new SalesOrderItemList();
            updatedOrder.salesOrder.orderStatus = SalesOrderOrderStatus._pendingFulfillment;
            updatedOrder.salesOrder.orderStatusSpecified = true;

            List<SalesOrderLine> newItemList = new List<SalesOrderLine>();
            newItemList = createNewSalesOrderLines(this.salesOrder.itemList);
            updatedOrder.addItemList(newItemList);

            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            WriteResponse writeResponse = service.update(updatedOrder.salesOrder);
            if (writeResponse.status.isSuccess == true)
            {
                Console.WriteLine("Approved Sales Order success");
            }
            if (writeResponse.status.isSuccess == false)
            {
                Console.WriteLine(writeResponse.status.statusDetail[0].message);
            }

            return this.salesOrder.tranId;
        }
    }
}