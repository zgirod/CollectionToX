using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollectionToX.Attributes
{

    public enum HorizontalAlignment
    {
        General,
        Left,
        Right,
        Center,
        Fill

    }

    public enum VerticalAlignment
    {
        Bottom,
        Center,
        Distributed,
        Justify,
        Top
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelPropertyStyleExportAttribute : Attribute
    {

        public string HeaderName { get; set; }

        public int? NumberFormatId { get; set; }
        public string NumberFormatCustom { get; set; }

        public HorizontalAlignment HorizontalAlignment { get; set; }
        public VerticalAlignment VerticalAlignment { get; set; }

        internal XLAlignmentHorizontalValues ExcelHorizontalAlignment
        {
            get
            {

                if (HorizontalAlignment == HorizontalAlignment.General) return XLAlignmentHorizontalValues.General;
                else if (HorizontalAlignment == HorizontalAlignment.Left) return XLAlignmentHorizontalValues.Left;
                else if (HorizontalAlignment == HorizontalAlignment.Right) return XLAlignmentHorizontalValues.Right;
                else if (HorizontalAlignment == HorizontalAlignment.Center) return XLAlignmentHorizontalValues.Center;
                else if (HorizontalAlignment == HorizontalAlignment.Fill) return XLAlignmentHorizontalValues.Fill;
                else return XLAlignmentHorizontalValues.General;

            }
        }

        internal XLAlignmentVerticalValues ExcelVerticalAlignment
        {
            get
            {

                if (VerticalAlignment == VerticalAlignment.Bottom) return XLAlignmentVerticalValues.Bottom;
                else if (VerticalAlignment == VerticalAlignment.Center) return XLAlignmentVerticalValues.Center;
                else if (VerticalAlignment == VerticalAlignment.Distributed) return XLAlignmentVerticalValues.Distributed;
                else if (VerticalAlignment == VerticalAlignment.Justify) return XLAlignmentVerticalValues.Justify;
                else if (VerticalAlignment == VerticalAlignment.Top) return XLAlignmentVerticalValues.Top;
                else return XLAlignmentVerticalValues.Bottom;

            }
        }

        public ExcelPropertyStyleExportAttribute()
        {
            HeaderName = null;
            NumberFormatId = null;
            NumberFormatCustom = "General";
            HorizontalAlignment = HorizontalAlignment.General;
            VerticalAlignment = VerticalAlignment.Bottom;
        }

    }
}
