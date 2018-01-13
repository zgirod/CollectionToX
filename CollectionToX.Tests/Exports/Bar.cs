using CollectionToX.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollectionToX.Tests.Exports
{
    public class Bar
    {

        [ExcelPropertyStyleExport(HeaderName = "Text Test", NumberFormatCustom = "@")]
        public string Text { get; set; }

    }
}
