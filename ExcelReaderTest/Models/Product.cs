using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderTest.Models
{
    public class Product
    {
        [Description("名稱")]
        public string Name { get; set; }
        [Description("庫存")]
        public int Inventory { get; set; }
        [Description("單位")]
        public string Unit { get; set; }
        [Description("廠商")]
        public string Vender { get; set; }
        [Description("價格")]
        public int Price { get; set; }
        [Description("利潤")]
        public int Profit { get; set; }
    }
}
