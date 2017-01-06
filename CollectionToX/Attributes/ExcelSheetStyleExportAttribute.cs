using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollectionToX.Attributes
{

    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelSheetStyleExportAttribute : Attribute
    {

        public string FontFamily { get; set; }
        public bool BoldHeader { get; set; }
        public int FontSizeHeader { get; set; }
        public int FontSizeData { get; set; }

        public ExcelSheetStyleExportAttribute()
        {
            FontFamily = "Calibri";
            BoldHeader = true;
            FontSizeHeader = 11;
            FontSizeData = 11;
        }

    }
}
