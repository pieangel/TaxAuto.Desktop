using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaxAuto.Desktop.Models
{
    public class PurchaseItemDto
    {
        public string? ItemNameRaw { get; set; }

        public int Quantity { get; set; }

        public int UnitPrice { get; set; }

        public int Amount { get; set; }
    }
}
