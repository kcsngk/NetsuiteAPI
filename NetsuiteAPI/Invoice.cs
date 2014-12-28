using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetsuiteAPI.com.netsuite.webservices;
using System.Net;

namespace NetsuiteAPI
{
    public class Invoice
    {
        public com.netsuite.webservices.Invoice invoice {get;set;}

        public Invoice()
        {
        }

        public Invoice(SalesOrder salesOrder)
        {
            this.invoice = new com.netsuite.webservices.Invoice();

            RecordRef salesOrderID = new RecordRef();
            salesOrderID.type = RecordType.salesOrder;
            salesOrderID.typeSpecified = true;

            
            salesOrderID.internalId = salesOrder.salesOrder.internalId;

            this.invoice.createdFrom = salesOrderID;
            
            this.invoice.billAddress = salesOrder.salesOrder.billAddress;
            //this.invoice.shipAddress = salesOrder.salesOrder.shipAddress;
            this.invoice.entity = salesOrder.salesOrder.entity;
            this.invoice.subTotal = salesOrder.salesOrder.subTotal;

            InvoiceItemList invoiceItemList = new InvoiceItemList();
            this.invoice.itemList = addItemList(salesOrder.salesOrder.itemList);

            this.invoice.memo = salesOrder.salesOrder.memo;
            this.invoice.location = salesOrder.salesOrder.location;
            //this.invoice.fob = salesOrder.salesOrder.fob;
            
        }

        public InvoiceItemList addItemList( SalesOrderItemList salesOrderLines)
        {
            InvoiceItemList invItemList = new InvoiceItemList();

            int i = 0;
            InvoiceItem[] Items = new InvoiceItem[salesOrderLines.item.Count()];

            foreach (var itemLine in salesOrderLines.item)
            {
                Items[i] = new InvoiceItem();
                RecordRef item = new RecordRef();
                item.type = RecordType.inventoryItem;
                item.typeSpecified = true;
                // item.internalId = "1229";
                item.internalId = salesOrderLines.item[i].item.internalId;
                Items[i].item = item;

                RecordRef prLevel = new RecordRef();
                prLevel.type = RecordType.priceLevel;
                prLevel.internalId = "-1";
                prLevel.typeSpecified = true;

                Items[i].price = prLevel;
                Items[i].rate = Convert.ToString(salesOrderLines.item[i].rate);
                Items[i].quantity = salesOrderLines.item[i].quantity;
                Items[i].quantitySpecified = true;
                i++;
            }
            invItemList.item = Items;
            return invItemList;
        }

        public void createInvoice()
        {
            NetSuiteService service = new NetSuiteService();
            service.CookieContainer = new CookieContainer();
            NetsuiteUser user = new NetsuiteUser("3451682", "kevin.ng@tridentcase.com", "1026", "tridenT168");
            Passport passport = user.prepare(user);
            Status status = service.login(passport).status;

            WriteResponse writeResponse = service.add(this.invoice);
            if (writeResponse.status.isSuccess == true)
            {

                Console.WriteLine("Create Invoice success");

            }
            if (writeResponse.status.isSuccess == false)
            {
                Console.WriteLine(writeResponse.status.statusDetail[0].message);
            }
            
        }

        
    }
}
