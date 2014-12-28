using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NetsuiteAPI
{
    public class ItemFulfillmentLine
    {
        public string InternalID { get; set; }
        public double QuantityRequested { get; set; }
        public bool delete { get; set; }
        public Item item { get; set; }
        public Location location { get; set; }
        

        public ItemFulfillmentLine()
        {
        }

        public ItemFulfillmentLine(string itemName, double quantity , bool flag)
        {
            item = new Item(itemName);
            item.addItemRecord();
            this.QuantityRequested = quantity;
            this.delete = flag;
        }


        public ItemFulfillmentLine(string itemName, double quantity, bool flag,Location loc)
        {
            item = new Item(itemName);
            item.addItemRecord();
            this.QuantityRequested = quantity;
            this.delete = flag;
            this.location = loc;
        }
    }
}
