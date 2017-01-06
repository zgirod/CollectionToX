using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollectionToX.Helpers
{
    public static class ExcelSheetNameHelper
    {

        public static string CleanExcelSheetName(this string sheetName)
        {

            if (string.IsNullOrWhiteSpace(sheetName))
                throw new Exception("Sheet name cannot be null.");

            sheetName = sheetName.Trim();
            if (sheetName.ToLower() == "history")
                throw new Exception("Sheet name cannot be 'History'.");

            char[] separators = new char[] { '\\', '/', '*', '?', ':', '[',']' };
            string[] temp = sheetName.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            sheetName = String.Join("", temp);

            if (sheetName.Length > 31)
                sheetName = sheetName.Substring(0, 31);

            return sheetName;

        }

    }
}
