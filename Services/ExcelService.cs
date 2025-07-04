using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using ClosedXML.Excel;
using DemoMVCProject1.Models;
using OfficeOpenXml;

namespace DemoMVCProject1.Services
{
    public static class ExcelService
    {
        public static List<string> GetExcelHeaders(string filePath)
        {
            var headers = new List<string>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();
                var firstRow = worksheet.FirstRowUsed();

                foreach (var cell in firstRow.CellsUsed())
                {
                    headers.Add(cell.GetString());
                }
            }

            return headers;
        }

        public static List<Customer> ParseExcel(string filePath, List<ExcelColumnMapping> mappings)
        {
            var customers = new List<Customer>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // First sheet
                var headerRow = worksheet.FirstRowUsed();
                var dataRows = worksheet.RowsUsed().Skip(1); // Skip header

                // Map: DB Column → Excel Column Index
                var columnIndexes = mappings.ToDictionary(
                    m => m.DbColumn,
                    m => headerRow.Cells().FirstOrDefault(c => c.GetString().Trim() == m.ExcelColumn.Trim())?.Address.ColumnNumber ?? -1
                );

                // Ensure all mappings are valid
                if (columnIndexes.Values.Any(idx => idx == -1))
                    throw new Exception("One or more Excel columns could not be found based on mapping.");

                foreach (var row in dataRows)
                {
                    var customer = new Customer();

                    foreach (var map in mappings)
                    {
                        int columnIndex = columnIndexes[map.DbColumn];
                        var cellValue = row.Cell(columnIndex).GetString();

                        PropertyInfo prop = typeof(Customer).GetProperty(map.DbColumn);
                        if (prop != null && prop.CanWrite)
                        {
                            try
                            {
                                object convertedValue = Convert.ChangeType(cellValue, prop.PropertyType);
                                prop.SetValue(customer, convertedValue);
                            }
                            catch
                            {
                                // Log or ignore if data format is invalid
                            }
                        }
                    }

                    customers.Add(customer);
                }
            }

            return customers;
        }
    }

}