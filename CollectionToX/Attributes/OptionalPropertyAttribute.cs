using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollectionToX.Attributes
{

    [AttributeUsage(AttributeTargets.Property)]
    public class OptionalPropertyAttribute : Attribute
    {

        public string OptionalKey { get; set; }
        public bool ShowOnMatch { get; set; }

        public OptionalPropertyAttribute()
        {
        }

    }

}
