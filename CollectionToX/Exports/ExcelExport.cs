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

        }

        private IXLWorksheet _worksheet { get; set; }
        private int _rowNumber { get; set; }

        public ExcelExport AddCollectionToExcel<T>(IEnumerable<T> data, string sheetName)
        {

            //build the excel workbook
            _rowNumber = 1;
            BuildExcelWorkbook(data, typeof(T), sheetName);
            return this;

        }

        public MemoryStream SaveAsMemoryStream()
        {

            //return the workbook as a stream
            using (MemoryStream stream = new MemoryStream())
            {
                _workbook.SaveAs(stream);
                return stream;
            }

        }

        public void SaveAsExcelFile(string filepath)
        {

            //save the excel
            _workbook.SaveAs(filepath);

        }

        private void BuildExcelWorkbook<T>(IEnumerable<T> data, Type type, string sheetName)
        {

            _worksheet = _workbook.Worksheets.Add(sheetName.CleanExcelSheetName());
            var excelSheetStyleExportAttribute = (ExcelSheetStyleExportAttribute)type.GetCustomAttributes(true).FirstOrDefault(x => x is ExcelSheetStyleExportAttribute) ?? new ExcelSheetStyleExportAttribute();
            var propertyList = ObjectToPropertList.GetPropertyList(type);

            //add in the header row
            AddHeaderRow(propertyList, excelSheetStyleExportAttribute);

            //add in the data
            AddDataRows(data, propertyList, excelSheetStyleExportAttribute);

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
                _worksheet.Cell(_rowNumber, col).Value = headerName;
                _worksheet.Cell(_rowNumber, col).Style.Font.FontName = excelSheetStyleExportAttribute.FontFamily;
                _worksheet.Cell(_rowNumber, col).Style.Font.FontSize = excelSheetStyleExportAttribute.FontSizeHeader;
                _worksheet.Cell(_rowNumber, col).Style.Font.Bold = excelSheetStyleExportAttribute.BoldHeader;

                //go to the next column
                col++;

            }

            //go to the next row
            _rowNumber++;

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
                    

                    //set the value
                    _worksheet.Cell(_rowNumber, col).Value = property.PropertyInfo.GetValue(item, null).ToString();

                    //set column based styles
                    _worksheet.Cell(_rowNumber, col).Style.NumberFormat.Format = excelPropertyStyleExportAttribute.NumberFormat;
                    _worksheet.Cell(_rowNumber, col).Style.Alignment.Vertical = excelPropertyStyleExportAttribute.ExcelVerticalAlignment;
                    _worksheet.Cell(_rowNumber, col).Style.Alignment.Horizontal = excelPropertyStyleExportAttribute.ExcelHorizontalAlignment;

                    //set the sheet based styles
                    _worksheet.Cell(_rowNumber, col).Style.Font.FontName = excelSheetStyleExportAttribute.FontFamily;
                    _worksheet.Cell(_rowNumber, col).Style.Font.FontSize = excelSheetStyleExportAttribute.FontSizeData;

                    //go to the next row
                    col++;

                }

                //go to the next row
                _rowNumber++;

            }

        }

    }

}
