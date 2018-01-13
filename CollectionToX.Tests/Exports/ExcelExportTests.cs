using CollectionToX.Attributes;
using CollectionToX.Exports;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CollectionToX.Tests.Exports
{

    public class ExcelExportTests
    {

        [Test]
        public void ExcelExport_Save5RowCollection()
        {

            var filepath = $"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}/foo.xlsx";
            if (File.Exists(filepath)) File.Delete(filepath);

            var data = new List<Foo>()
                {
                    new Foo() { Name = "Marshall", City = "Paw", State = "MN" },
                    new Foo() { Name = "Carlos", City = "Patrol", State = "MD" },
                    new Foo() { Name = "Carlos", City = "Paw", State = "ND", NullableInt = 19 }
                };

            var export = new ExcelExport();
            export.AddCollectionToExcel<Foo>(data, "first");
            export.AddCollectionToExcel<Foo>(data, "second");

            export.AddBlankRowsToWorksheet("first", 10);
            export.AddCollectionToExistingSheet<Foo>(data, "first");

            export.SaveAsExcelFile(filepath);

            //TODO: Assert something

        }

        [Test]
        public void ExcelExport_SaveMemoryStream()
        {

            var data = new List<Foo>()
            {
                new Foo() { Name = "Marshall", City = "Paw", State = "MN" },
                new Foo() { Name = "Carlos", City = "Patrol", State = "MD" },
                new Foo() { Name = "Carlos", City = "Paw", State = "ND", NullableInt = 19 }
            };

            var export = new ExcelExport();
            export.AddCollectionToExcel<Foo>(data, "first");
            export.AddCollectionToExcel<Foo>(data, "second");
            var ms = export.SaveAsMemoryStream();

        }

        [Test]
        public void ExcelExport_TextTest()
        {


            var filepath = $"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}/text.xlsx";
            var data = new List<Bar>()
            {
                new Bar() { Text = "121196" },
                new Bar() { Text = "00121196" },
                new Bar() { Text = "1211E6" },
                new Bar() { Text = "1211e6" },
                new Bar() { Text = "00121AAA" }
            };

            var export = new ExcelExport();
            export.AddCollectionToExcel<Bar>(data, "first");
            export.AddCollectionToExcel<Bar>(data, "second");
            export.SaveAsExcelFile(filepath);

        }

    }
}
