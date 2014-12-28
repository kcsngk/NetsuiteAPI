using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NetsuiteAPI
{
    public class SalesOrderLine
    {
        public string InternalID { get; set; }
        public double QuantityRequested { get; set; }
        public decimal UnitPrice { get; set; }
        public bool delete { get; set; }
        public Item item { get; set; }

        public SalesOrderLine()
        {
        }

        public SalesOrderLine(string itemName, double quantity ,decimal price, bool flag)
        {
            item = new Item(itemName);
            item.addItemRecord();
            this.QuantityRequested = quantity;
            this.UnitPrice = price;
            this.delete = flag;
        }

     
    }
}
