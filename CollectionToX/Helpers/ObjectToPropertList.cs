using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web;
using System.Threading.Tasks;
using CollectionToX.Attributes;

namespace CollectionToX.Helpers
{

    public static class ObjectToPropertList
    {

        public static List<PropertyListItem> GetPropertyList(Type type)
        {

            //get the all prperties we want returned
            var properties = type.GetProperties()
                .Select(x => new
                {
                    PropertyInfo = x,
                    CustomAttributes = x.GetCustomAttributes(true)
                })
                .Where(x => x.CustomAttributes.Any(attribute => attribute is IgnorePropertyExportAttribute) == false)
                .ToList();

            var propertyListItems = new List<PropertyListItem>();

            //add in all the properties with custom sort order
            foreach(var property in properties
                .Where(x => x.CustomAttributes.Any(attribute => attribute is SortOrderExportAttribute))
                .OrderBy(x => ((SortOrderExportAttribute)x.CustomAttributes.First(attribute => attribute is SortOrderExportAttribute)).SortOrder)
                .ToList())
            {
                propertyListItems.Add(new PropertyListItem() { PropertyInfo = property.PropertyInfo, CustomAttributes = property.CustomAttributes });
            }

            //add in all the properties with no custom sort order
            foreach (var property in properties
                .Where(x => x.CustomAttributes.Any(attribute => attribute is SortOrderExportAttribute) == false)
                .ToList())
            {
                propertyListItems.Add(new PropertyListItem() { PropertyInfo = property.PropertyInfo, CustomAttributes = property.CustomAttributes });
            }

            return propertyListItems;

        }

    }
}
