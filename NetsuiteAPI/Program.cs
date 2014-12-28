using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetsuiteAPI.com.netsuite.webservices;
using System.Net;

namespace NetsuiteAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            /********************** Create new Bins/Delete Bin
            BinManagement binCreation = new BinManagement();
            
            List<InventoryBin> listOfBins = new List<InventoryBin>();
            listOfBins = binCreation.GetBinsFromExcel(args[0]);

            InventoryBin ingramMalaysia = new InventoryBin("Consignment","IngramMalaysia");

            foreach (var bin in listOfBins)
            {
               binCreation.createNewBin(ingramMalaysia);
            }
            
            //InventoryBin testBin = new InventoryBin("Amazon - US","TestBin");
            //binCreation.DeleteBin(testBin);
            
             */
            
            
            /**********************  Add bins to items
            ItemBinUpdate itemBinUpdate = new ItemBinUpdate();
            List<InventoryBin> listOfItemBins = new List<InventoryBin>();
            listOfItemBins = itemBinUpdate.GetItemBinsFromExcel(args[0]);

            foreach (var itemBin in listOfItemBins)
            {
                itemBinUpdate.addBin(itemBin);
            }
            */

            
            
            //********************** Create new Sales Order
            /*
            bool noDelete = false;
            SalesOrderInfo salesInfo = new SalesOrderInfo("test address", "test address2", "test order", "FOB");
         
            SalesOrder testOrder = new SalesOrder(salesInfo);

            List<SalesOrderLine> orderItems = new List<SalesOrderLine>();
            orderItems.Add(new SalesOrderLine("AG-3GS-BK",1,15,noDelete));
            orderItems.Add(new SalesOrderLine("AG-3GS-BL", 1, 12,noDelete));
            orderItems.Add(new SalesOrderLine("AG-3GS-PK", 1, 10,noDelete));

            testOrder.addItemList(orderItems);
            testOrder.setCustomer("Test Celigo Company");
            testOrder.placeSalesOrder();
            */
            
            SalesOrder findOrder = new SalesOrder();
            findOrder.salesOrderSearch("S39853");
            /*
            foreach (var field in findOrder.salesOrder.customFieldList)
            {
                string intID = field.internalId;
                if(intID == "148")
                {
                    Console.WriteLine(((com.netsuite.webservices.SelectCustomFieldRef)field).value.name);
                }
            }
            */
            
            //string test = findOrder.approveOrder();


            /*********************** Edit Sales Order
            
            //bool delete = true;
            bool noDelete = false;
            List<SalesOrderLine> editList = new List<SalesOrderLine>();
            editList.Add(new SalesOrderLine("AG-ALOTC7-TG000", 1, 15, noDelete));
           
            
            findOrder.editOrder(editList);
            */ 
            
            //****************** Create new Invoice
            //Invoice testInvoice = new Invoice(findOrder);
            //testInvoice.createInvoice();
            
              
            //***************** Create new Item Fulfillment  
            ItemFulfillment testIF = new ItemFulfillment(findOrder);
            testIF.createItemFulfillment();

            //*
            //ItemFulfillment findFulfillment = new ItemFulfillment();
            //findFulfillment.itemFulfillmentSearch("18011");

           // bool delete = false;
            //List<ItemFulfillmentLine> editLines = new List<ItemFulfillmentLine>();
            //editLines.Add(new ItemFulfillmentLine("AG-3GS-BK", 3,delete));

            //findFulfillment.editFulfillment(editLines);
            //*/ 

           
            

            //testOrder.deleteOrder();

            //Item testItem = new Item("AG-3GS-BK");
           // testItem.addItemRecord();

            //string test = findFulfillment.markAsShipped();

           //Console.WriteLine(test);
           
            Console.Read();
         
        }
    }
}
