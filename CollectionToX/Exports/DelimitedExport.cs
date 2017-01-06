using CollectionToX.Attributes;
using CollectionToX.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollectionToX.Exports
{

    public enum Delimiter
    {
        COMMA,
        TAB,
        PIPE
    }

    public class DelimitedExport
    {

        private StringBuilder _stringBuilder { get; set; }
        private char _delimiter { get; set; }

        public DelimitedExport(Delimiter delimiter)
        {
            _stringBuilder = new StringBuilder();

            if (delimiter == Delimiter.COMMA)
                _delimiter = ',';
            else if (delimiter == Delimiter.TAB)
                _delimiter = '\t';
            else if (delimiter == Delimiter.PIPE)
                _delimiter = '|';

        }

        public DelimitedExport AddLine(string[] line)
        {

            var first = true;
            var sbRow = new StringBuilder();

            if (line != null)
            {
                foreach (var item in line)
                {

                    if (first == false)
                        sbRow.Append(_delimiter);
                    sbRow.AppendFormat("\"{0}\"", item);
                    first = false;

                }
            }

            //add data to the stream
            _stringBuilder.AppendLine(sbRow.ToString());
            return this;

        }

        public DelimitedExport AddCollectionToExcel<T>(IEnumerable<T> data)
        {

            //add data to the stream
            BuildDelimitedExport(data, typeof(T));
            return this;

        }

        public MemoryStream SaveAsMemoryStream()
        {

            //return the workbook as a stream
            using (MemoryStream stream = new MemoryStream())
            {
                var encoding = new ASCIIEncoding();
                stream.Write(encoding.GetBytes(_stringBuilder.ToString()), 0, _stringBuilder.Length);
                return stream;
            }

        }

        public void SaveAsFile(string filepath)
        {

            using (StreamWriter outputFile = new StreamWriter(filepath))
            {
                outputFile.WriteLine(_stringBuilder.ToString());
            }

        }

        private void BuildDelimitedExport<T>(IEnumerable<T> data, Type type)
        {

            var propertyList = ObjectToPropertList.GetPropertyList(type);

            //add in the header row
            AddHeaderRow(propertyList);

            //add in the data
            AddDataRows(data, propertyList);

        }

        private void AddHeaderRow(List<PropertyListItem> propertyList)
        {

            //for each property, add in the header
            var first = true;
            var sbRow = new StringBuilder();
            foreach (var property in propertyList)
            {

                if (first == false)
                    sbRow.Append(_delimiter);

                //get the properties for this attribute
                var excelPropertyStyleExportAttribute = (ExcelPropertyStyleExportAttribute)property.CustomAttributes.FirstOrDefault(attribute => attribute is ExcelPropertyStyleExportAttribute) ?? new ExcelPropertyStyleExportAttribute();

                //set the header name
                var headerName = property.PropertyInfo.Name;
                if (string.IsNullOrWhiteSpace(excelPropertyStyleExportAttribute.HeaderName) == false)
                    headerName = excelPropertyStyleExportAttribute.HeaderName;
                sbRow.AppendFormat("\"{0}\"", headerName);

                //go to the next column
                first = false;

            }

            //add the header row
            _stringBuilder.AppendLine(sbRow.ToString());

        }

        private void AddDataRows<T>(IEnumerable<T> data, List<PropertyListItem> propertyList)
        {

            //for each item
            foreach (var item in data)
            {

                var first = true;
                var sbRow = new StringBuilder();
                foreach (var property in propertyList)
                {

                    if (first == false)
                        sbRow.Append(_delimiter);

                    sbRow.AppendFormat("\"{0}\"", property.PropertyInfo.GetValue(item, null).ToString());
                    first = false;

                }

                //append the full line
                _stringBuilder.AppendLine(sbRow.ToString());

            }

        }

    }

}

