using CollectionToX.Exports;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace CollectionToX
{
    public static class CollectionToXExtensions
    {

        public static MemoryStream ToExcelStream<T>(this IEnumerable<T> data) where T : class
        {
            return ToExcelStream<T>(data, typeof(T).Name);
        }

        public static MemoryStream ToExcelStream<T>(this IEnumerable<T> data, string sheetName) where T : class
        {
            var export = new ExcelExport();
            return export.AddCollectionToExcel<T>(data, sheetName)
                .SaveAsMemoryStream();
        }

        public static void ToExcelFile<T>(this IEnumerable<T> data, string filepath) where T : class
        {
            ToExcelFile<T>(data, typeof(T).Name, filepath);
        }

        public static void ToExcelFile<T>(this IEnumerable<T> data, string sheetName, string filepath) where T : class
        {
            var export = new ExcelExport();
            export.AddCollectionToExcel<T>(data, sheetName)
                .SaveAsExcelFile(filepath);
        }

        public static HttpResponseMessage ToExcelHttpResponseMessage<T>(this IEnumerable<T> data, string sheetName, string fileName) where T : class
        {
            var export = new ExcelExport();
            var ms = export.AddCollectionToExcel<T>(data, sheetName).SaveAsMemoryStream();
            ms.Seek(0, SeekOrigin.Begin);
            var httpResponse = new HttpResponseMessage();
            httpResponse.Content = new StreamContent(ms);
            httpResponse.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = fileName
            };
            httpResponse.Content.Headers.Add("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            return httpResponse;
        }

        public static HttpResponseMessage ToDelimitedHttpResponseMessage<T>(this IEnumerable<T> data, Delimiter delimiter, string fileName) where T : class
        {
            var export = new DelimitedExport(delimiter);
            var ms = export.AddCollectionToExcel<T>(data).SaveAsMemoryStream();
            ms.Seek(0, SeekOrigin.Begin);
            var httpResponse = new HttpResponseMessage();
            httpResponse.Content = new StreamContent(ms);
            httpResponse.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = fileName
            };

            if (delimiter == Delimiter.COMMA) httpResponse.Content.Headers.Add("Content-Type", "text/csv");
            else if (delimiter == Delimiter.TAB) httpResponse.Content.Headers.Add("Content-Type", "text/tab-separated-values");
            else if (delimiter == Delimiter.PIPE) httpResponse.Content.Headers.Add("Content-Type", "text/plain");

            return httpResponse;
        }

    }

}
