using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollectionToX.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class IgnorePropertyExportAttribute : Attribute
    {
    }
}
