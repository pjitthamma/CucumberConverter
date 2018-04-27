using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Collections.Generic;
using CucumberConverter.Common.model;

namespace CucumberConverter.Common
{
    public class ExcelEPPluse
    {
        #region Properties
        private const int EXCEL_START_ROW = 5;
        #endregion

        public ExcelEPPluse(){}

        public TCExcelDto Read(string path, int pageNumber)
        {
            // Prepare excel
            var excel = new ExcelPackage(new FileInfo(path));
            ExcelWorksheet workSheet = excel.Workbook.Worksheets[pageNumber];

            // Read data
            var result = new TCExcelDto();
            for (int row = EXCEL_START_ROW; row <= workSheet.Dimension.End.Row; row++)
            {
                Dictionary<int, Object> excelDictionary = new Dictionary<int, Object>();
                for (int column = 1; column <= workSheet.Dimension.End.Column; column++)
                {
                    Object value = workSheet.Cells[row, column].Value;
                    excelDictionary.Add(column, value);
                }
                if (excelDictionary.Any()) { result.TCList.Add(excelDictionary);  }
            }

            return result;
        }

        public void Write(string excelLocation, int pageNumber, Dictionary<string, string> datas)
        {
            System.IO.FileInfo file = new System.IO.FileInfo(excelLocation);
            using (var excelFile = new ExcelPackage(file))
            {
                var worksheet = excelFile.Workbook.Worksheets[pageNumber];

                foreach (var data in datas)
                {
                    worksheet.Cells[data.Key].LoadFromText(data.Value);
                }
                excelFile.Save();
                excelFile.Dispose();
            } 
        }



        #region Helper method
        public Dictionary<string, string> BuildWriteData(int index, List<string> columns, List<string> values)
        {
            var excelDatas = new Dictionary<string, string>();

            for(var i=0; i<columns.Count; i++)
            {
                excelDatas.Add(columns[i] + (index+EXCEL_START_ROW), values[i]);
            }

            return excelDatas;
        }
        #endregion
    }
}
