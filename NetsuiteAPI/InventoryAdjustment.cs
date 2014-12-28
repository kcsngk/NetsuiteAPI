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
    public class InventoryAdjustment
    {
        public NetsuiteAPI.com.netsuite.webservices.InventoryAdjustment invAdjustment{ get; set; }
        

        public InventoryAdjustment()
        {
        }

        public List<InventoryBin> GetInventoryBinsFromExcel(string workbook)
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
                    int quantity = Convert.ToInt16(excelRange.Cells[i, 2].Value2.ToString());
                    string binNumber = excelRange.Cells[i, 3].Value2.ToString();
                    string itemName = excelRange.Cells[i, 1].Value2.ToString();
                    // InventoryBin binToAdd = new InventoryBin(location,binNumber);
                    Item itemToAdd = new Item(itemName);
                    itemToAdd.addItemRecord();
                    itemBins.Add(new InventoryBin(quantity, binNumber, itemToAdd));
                }
            }

            return itemBins;

        }
        
        public void makeAdjustment(List<InventoryBin> itemsToAdjust)
        {
            invAdjustment = new com.netsuite.webservices.InventoryAdjustment();
            InventoryAdjustmentInventory[] invAdjustmentItemArray = new InventoryAdjustmentInventory[itemsToAdjust.Count];

        }

    }
}
