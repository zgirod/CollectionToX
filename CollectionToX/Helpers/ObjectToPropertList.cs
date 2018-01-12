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

        public static List<PropertyListItem> GetPropertyList(Type type, string optionalKey)
        {

            //get the all properties we want returned
            //we don't want properties with the ignore attribute
            //if there is an optional attribute the optional key has to be passed in or there is no optional attribute
            var properties = type.GetProperties()
                .Select(x => new
                {
                    PropertyInfo = x,
                    CustomAttributes = x.GetCustomAttributes(true)
                })
                .Where(x => 
                    x.CustomAttributes.Any(attribute => attribute is IgnorePropertyExportAttribute) == false
                    &&
                    (
                        x.CustomAttributes.Any(attribute => attribute is OptionalPropertyAttribute) == false
                        ||
                        (
                            (
                                x.CustomAttributes.Any(attribute => attribute is OptionalPropertyAttribute) == true
                                && optionalKey != null
                                && ((OptionalPropertyAttribute)x.CustomAttributes.First(attribute => attribute is OptionalPropertyAttribute)).OptionalKey == optionalKey
                                && ((OptionalPropertyAttribute)x.CustomAttributes.First(attribute => attribute is OptionalPropertyAttribute)).ShowOnMatch == true
                            )
                            ||
                            (
                                x.CustomAttributes.Any(attribute => attribute is OptionalPropertyAttribute) == true
                                && ((OptionalPropertyAttribute)x.CustomAttributes.First(attribute => attribute is OptionalPropertyAttribute)).OptionalKey != optionalKey
                                && ((OptionalPropertyAttribute)x.CustomAttributes.First(attribute => attribute is OptionalPropertyAttribute)).ShowOnMatch == false
                            )
                        )
                    )
                )
                .ToList();

            var propertyListItems = new List<PropertyListItem>();

            //add in all the properties with custom sort order
            foreach (var property in properties
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
