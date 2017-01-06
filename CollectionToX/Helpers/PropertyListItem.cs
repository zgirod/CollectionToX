using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CollectionToX.Helpers
{
    public class PropertyListItem
    {
        public PropertyInfo PropertyInfo { get; set; }
        public object[] CustomAttributes { get; set; }
    }
}
