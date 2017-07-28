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
    public class DelimitedExportTests
    {

        [Test]
        public void DelimitedExport_SaveCollectionAsCSV()
        {

            var filepath = $"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}/foo.csv";
            if (File.Exists(filepath)) File.Delete(filepath);

            var data = new List<Foo>()
            {
                new Foo() { Name = "Marshall", City = "Paw", State = "MN", FoundingDate = DateTime.Parse("7/1/2017") },
                new Foo() { Name = "Carlos", City = "Patrol", State = "MD", FoundingDate = DateTime.Parse("7/1/2017 10:10:10") },
                new Foo() { Name = "Carlos", City = "Patrol", State = "MD", NullableInt = 19 }
            };

            var export = new DelimitedExport(Delimiter.COMMA);
            export.AddCollection<Foo>(data);
            export.SaveAsFile(filepath);

            //TODO: Assert something

        }

        [Test]
        public void DelimitedExport_SaveCollectionAsPipeDelimited()
        {

            var filepath = $"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}/foo_pipe.txt";
            if (File.Exists(filepath)) File.Delete(filepath);

            var data = new List<Foo>()
            {
                new Foo() { Name = "Marshall", City = "Paw", State = "MN" },
                new Foo() { Name = "Carlos", City = "Patrol", State = "MD" }
            };

            var export = new DelimitedExport(Delimiter.PIPE);
            export.AddLine(new string[] { "This is the Title", DateTime.Parse("5/5/2000").ToShortDateString() });
            export.AddCollection<Foo>(data);
            export.SaveAsFile(filepath);

            //TODO: Assert something

        }

    }
}
