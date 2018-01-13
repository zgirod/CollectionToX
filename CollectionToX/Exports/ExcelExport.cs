using CollectionToX.Helpers;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CollectionToX.Attributes;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CollectionToX.Exports
{
    public class ExcelExport
    {

        private XLWorkbook _workbook { get; set; }

        public ExcelExport()
        {

            //setup the excel sheet
            _workbook = new XLWorkbook();
            _worksheetRowNumbers = new Dictionary<string, int>();

        }

        private IXLWorksheet _worksheet { get; set; }
        private Dictionary<string, int> _worksheetRowNumbers;

        public ExcelExport AddCollectionToExcel<T>(IEnumerable<T> data, string sheetName, string optionalKey)
        {

            //build the excel workbook
            _worksheetRowNumbers.Add(sheetName.CleanExcelSheetName(), 1);
            BuildExcelWorksheet(data, typeof(T), sheetName, optionalKey);
            return this;

        }

        public ExcelExport AddCollectionToExcel<T>(IEnumerable<T> data, string sheetName)
        {

            return AddCollectionToExcel<T>(data, sheetName, null);

        }

        public ExcelExport AddCollectionToExistingSheet<T>(IEnumerable<T> data, string sheetName, string optionalKey)
        {

            AddToExcelWorksheet(data, typeof(T), sheetName, optionalKey);
            return this;

        }

        public ExcelExport AddCollectionToExistingSheet<T>(IEnumerable<T> data, string sheetName)
        {

            AddToExcelWorksheet(data, typeof(T), sheetName, null);
            return this;

        }

        public ExcelExport AddBlankRowsToWorksheet(string sheetName, int numberOfRows)
        {

            for(int i = 1; i <= numberOfRows; i++)
                _worksheetRowNumbers[sheetName.CleanExcelSheetName()] = _worksheetRowNumbers[sheetName.CleanExcelSheetName()] + 1;
            return this;

        }

        public MemoryStream SaveAsMemoryStream()
        {

            //return the workbook as a stream
            using (MemoryStream stream = new MemoryStream())
            {
                _workbook.SaveAs(stream);
                stream.Seek(0, SeekOrigin.Begin);
                var copyStream = new MemoryStream();
                stream.CopyTo(copyStream);
                copyStream.Seek(0, SeekOrigin.Begin);
                return copyStream;
            }

        }

        public void SaveAsExcelFile(string filepath)
        {

            //save the excel
            _workbook.SaveAs(filepath);

        }

        private void BuildExcelWorksheet<T>(IEnumerable<T> data, Type type, string sheetName, string optionalKey)
        {

            _worksheet = _workbook.Worksheets.Add(sheetName.CleanExcelSheetName());
            var excelSheetStyleExportAttribute = (ExcelSheetStyleExportAttribute)type.GetCustomAttributes(true).FirstOrDefault(x => x is ExcelSheetStyleExportAttribute) ?? new ExcelSheetStyleExportAttribute();
            var propertyList = ObjectToPropertList.GetPropertyList(type, optionalKey);

            //add in the header row
            AddHeaderRow(propertyList, excelSheetStyleExportAttribute);

            //add in the data
            AddDataRows(data, propertyList, excelSheetStyleExportAttribute);

            //adjust each column to the contents
            for (int col = 1; col <= propertyList.Count(); col++)
                _worksheet.Column(col).AdjustToContents();

        }

        private void BuildExcelWorksheet<T>(IEnumerable<T> data, Type type, string sheetName)
        {

            BuildExcelWorksheet<T>(data, type, sheetName, null);

        }

        private void AddToExcelWorksheet<T>(IEnumerable<T> data, Type type, string sheetName, string optionalKey)
        {

            //get the existing excel
            var cleanName = sheetName.CleanExcelSheetName();
            _worksheet = _workbook.Worksheets.Single(x => x.Name == cleanName);
            var excelSheetStyleExportAttribute = (ExcelSheetStyleExportAttribute)type.GetCustomAttributes(true).FirstOrDefault(x => x is ExcelSheetStyleExportAttribute) ?? new ExcelSheetStyleExportAttribute();
            var propertyList = ObjectToPropertList.GetPropertyList(type, optionalKey);

            //add in the data
            AddDataRows(data, propertyList, excelSheetStyleExportAttribute);

            //adjust each column to the contents
            for (int col = 1; col <= propertyList.Count(); col++)
                _worksheet.Column(col).AdjustToContents();

        }

        private void AddHeaderRow(List<PropertyListItem> propertyList, ExcelSheetStyleExportAttribute excelSheetStyleExportAttribute)
        {

            //for each property, add in the header
            var col = 1;
            foreach (var property in propertyList)
            {

                //get the properties for this attribute
                var excelPropertyStyleExportAttribute = (ExcelPropertyStyleExportAttribute)property.CustomAttributes.FirstOrDefault(attribute => attribute is ExcelPropertyStyleExportAttribute) ?? new ExcelPropertyStyleExportAttribute();

                //set the header name
                var headerName = property.PropertyInfo.Name;
                if (string.IsNullOrWhiteSpace(excelPropertyStyleExportAttribute.HeaderName) == false)
                    headerName = excelPropertyStyleExportAttribute.HeaderName;
                _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Value = headerName;
                _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Style.Font.FontName = excelSheetStyleExportAttribute.FontFamily;
                _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Style.Font.FontSize = excelSheetStyleExportAttribute.FontSizeHeader;
                _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Style.Font.Bold = excelSheetStyleExportAttribute.BoldHeader;

                //go to the next column
                col++;

            }

            //go to the next row
            _worksheetRowNumbers[_worksheet.Name] = _worksheetRowNumbers[_worksheet.Name] + 1;

        }

        private void AddDataRows<T>(IEnumerable<T> data, List<PropertyListItem> propertyList, ExcelSheetStyleExportAttribute excelSheetStyleExportAttribute)
        {

            var attributeStyleDictionary = new Dictionary<string, ExcelPropertyStyleExportAttribute>();

            //for each item
            foreach (var item in data)
            {

                var col = 1;
                foreach (var property in propertyList)
                {

                    ExcelPropertyStyleExportAttribute excelPropertyStyleExportAttribute = null;
                    if (attributeStyleDictionary.ContainsKey(property.PropertyInfo.Name))
                    {
                        excelPropertyStyleExportAttribute = attributeStyleDictionary[property.PropertyInfo.Name];
                    }
                    else
                    {

                        //get the properties for this attribute
                        excelPropertyStyleExportAttribute = (ExcelPropertyStyleExportAttribute)property.CustomAttributes.FirstOrDefault(attribute => attribute is ExcelPropertyStyleExportAttribute) ?? new ExcelPropertyStyleExportAttribute();
                        attributeStyleDictionary.Add(property.PropertyInfo.Name, excelPropertyStyleExportAttribute);

                    }

                    //set column based styles
                    if (excelPropertyStyleExportAttribute.NumberFormatId.HasValue == true)
                    {
                        _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Style.NumberFormat.NumberFormatId = excelPropertyStyleExportAttribute.NumberFormatId.Value;
                    }
                    else
                    {
                        _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Style.NumberFormat.Format = excelPropertyStyleExportAttribute.NumberFormatCustom;
                    }
                    
                    _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Style.Alignment.Vertical = excelPropertyStyleExportAttribute.ExcelVerticalAlignment;
                    _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Style.Alignment.Horizontal = excelPropertyStyleExportAttribute.ExcelHorizontalAlignment;

                    //set the sheet based styles
                    _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Style.Font.FontName = excelSheetStyleExportAttribute.FontFamily;
                    _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Style.Font.FontSize = excelSheetStyleExportAttribute.FontSizeData;

                    //set the value
                    var obj = property.PropertyInfo.GetValue(item, null);
                    if (obj != null)
                        _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Value = obj.ToString();
                    else
                        _worksheet.Cell(_worksheetRowNumbers[_worksheet.Name], col).Value = null;

                    //go to the next row
                    col++;

                }

                //go to the next row
                _worksheetRowNumbers[_worksheet.Name] = _worksheetRowNumbers[_worksheet.Name] + 1;

            }

        }

    }

}
