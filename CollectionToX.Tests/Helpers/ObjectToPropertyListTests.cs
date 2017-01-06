using CollectionToX.Attributes;
using CollectionToX.Helpers;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollectionToX.Tests.Helpers
{

    public class NoAttributes
    {
        public int First { get; set; }
        public string Second { get; set; }
    }

    public class IgnoreAttributes
    {
        [IgnorePropertyExport]
        public int First { get; set; }
        public string Second { get; set; }
    }

    public class IgnoreAndOrderAttributes
    {
        
        public int First { get; set; }
        [IgnorePropertyExport]
        public string Second { get; set; }
        [SortOrderExport(SortOrder = 2)]
        public decimal Third { get; set; }
        [SortOrderExport(SortOrder = 1)]
        public DateTime? Fourth { get; set; }
    }

    public class ObjectToPropertyListTests
    {

        [Test]
        public void ObjectToPropertyList_NoAttributesTest()
        {

            var propertyList = ObjectToPropertList.GetPropertyList(typeof(NoAttributes));
            Assert.That(propertyList.Count, Is.EqualTo(2));
            Assert.That(propertyList[0].PropertyInfo.Name, Is.EqualTo("First"));
            Assert.That(propertyList[1].PropertyInfo.Name, Is.EqualTo("Second"));
        }

        [Test]
        public void ObjectToPropertyList_IgnoreAttributesTest()
        {

            var propertyList = ObjectToPropertList.GetPropertyList(typeof(IgnoreAttributes));
            Assert.That(propertyList.Count, Is.EqualTo(1));
            Assert.That(propertyList[0].PropertyInfo.Name, Is.EqualTo("Second"));
        }

        [Test]
        public void ObjectToPropertyList_IgnoreAndOrderAttributesTest()
        {

            var propertyList = ObjectToPropertList.GetPropertyList(typeof(IgnoreAndOrderAttributes));
            Assert.That(propertyList.Count, Is.EqualTo(3));
            Assert.That(propertyList[0].PropertyInfo.Name, Is.EqualTo("Fourth"));
            Assert.That(propertyList[1].PropertyInfo.Name, Is.EqualTo("Third"));
            Assert.That(propertyList[2].PropertyInfo.Name, Is.EqualTo("First"));
        }


    }
}
