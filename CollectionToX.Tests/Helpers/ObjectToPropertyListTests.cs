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

    public class IgnoreOrderAndOptionalAttributes
    {

        public int First { get; set; }
        [IgnorePropertyExport]
        public string Second { get; set; }
        [SortOrderExport(SortOrder = 2)]
        public decimal Third { get; set; }
        [SortOrderExport(SortOrder = 1)]
        public int Fourth { get; set; }
        public string Fifth { get; set; }

        [SortOrderExport(SortOrder = 3)]
        [OptionalProperty(OptionalKey = "PASSPORT")]
        public double SixthOptional { get; set; }

        [SortOrderExport(SortOrder = 4)]
        public decimal Seven { get; set; }
    }

    public class ObjectToPropertyListTests
    {

        [Test]
        public void ObjectToPropertyList_NoAttributesTest()
        {

            var propertyList = ObjectToPropertList.GetPropertyList(typeof(NoAttributes), null);
            Assert.That(propertyList.Count, Is.EqualTo(2));
            Assert.That(propertyList[0].PropertyInfo.Name, Is.EqualTo("First"));
            Assert.That(propertyList[1].PropertyInfo.Name, Is.EqualTo("Second"));
        }

        [Test]
        public void ObjectToPropertyList_IgnoreAttributesTest()
        {

            var propertyList = ObjectToPropertList.GetPropertyList(typeof(IgnoreAttributes), null);
            Assert.That(propertyList.Count, Is.EqualTo(1));
            Assert.That(propertyList[0].PropertyInfo.Name, Is.EqualTo("Second"));
        }

        [Test]
        public void ObjectToPropertyList_IgnoreAndOrderAttributesTest()
        {

            var propertyList = ObjectToPropertList.GetPropertyList(typeof(IgnoreAndOrderAttributes), null);
            Assert.That(propertyList.Count, Is.EqualTo(3));
            Assert.That(propertyList[0].PropertyInfo.Name, Is.EqualTo("Fourth"));
            Assert.That(propertyList[1].PropertyInfo.Name, Is.EqualTo("Third"));
            Assert.That(propertyList[2].PropertyInfo.Name, Is.EqualTo("First"));
        }

        [Test]
        public void ObjectToPropertyList_IgnoreOrderAndOptionalExcludeNullAttributesTest()
        {
            var propertyList = ObjectToPropertList.GetPropertyList(typeof(IgnoreOrderAndOptionalAttributes), null);
            Assert.That(propertyList.Count, Is.EqualTo(5));
            Assert.That(propertyList[0].PropertyInfo.Name, Is.EqualTo("Fourth"));
            Assert.That(propertyList[1].PropertyInfo.Name, Is.EqualTo("Third"));
            Assert.That(propertyList[2].PropertyInfo.Name, Is.EqualTo("Seven"));
            Assert.That(propertyList[3].PropertyInfo.Name, Is.EqualTo("First"));
            Assert.That(propertyList[4].PropertyInfo.Name, Is.EqualTo("Fifth"));
        }

        [Test]
        public void ObjectToPropertyList_IgnoreOrderAndOptionalExcludeWrongKeyAttributesTest()
        {
            var propertyList = ObjectToPropertList.GetPropertyList(typeof(IgnoreOrderAndOptionalAttributes), "LICENSE");
            Assert.That(propertyList.Count, Is.EqualTo(5));
            Assert.That(propertyList[0].PropertyInfo.Name, Is.EqualTo("Fourth"));
            Assert.That(propertyList[1].PropertyInfo.Name, Is.EqualTo("Third"));
            Assert.That(propertyList[2].PropertyInfo.Name, Is.EqualTo("Seven"));
            Assert.That(propertyList[3].PropertyInfo.Name, Is.EqualTo("First"));
            Assert.That(propertyList[4].PropertyInfo.Name, Is.EqualTo("Fifth"));
        }

        [Test]
        public void ObjectToPropertyList_IgnoreOrderAndOptionalIncludeAttributesTest()
        {
            var propertyList = ObjectToPropertList.GetPropertyList(typeof(IgnoreOrderAndOptionalAttributes), "PASSPORT");
            Assert.That(propertyList.Count, Is.EqualTo(6));
            Assert.That(propertyList[0].PropertyInfo.Name, Is.EqualTo("Fourth"));
            Assert.That(propertyList[1].PropertyInfo.Name, Is.EqualTo("Third"));
            Assert.That(propertyList[2].PropertyInfo.Name, Is.EqualTo("SixthOptional"));
            Assert.That(propertyList[3].PropertyInfo.Name, Is.EqualTo("Seven"));
            Assert.That(propertyList[4].PropertyInfo.Name, Is.EqualTo("First"));
            Assert.That(propertyList[5].PropertyInfo.Name, Is.EqualTo("Fifth"));
        }


    }
}
