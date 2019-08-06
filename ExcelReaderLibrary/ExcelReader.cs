using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace ExcelReaderLibrary
{
    public class ExcelReader<T> : IExcelReader<T> where T : class
    {
        public List<T> ReadExcel(string filePath, bool useDescription = true)
        {
            var titleKeys = new List<Tuple<string, string>>();
            foreach (var prop in typeof(T).GetProperties())
            {
                var description = prop.GetCustomAttribute(typeof(DescriptionAttribute)) as DescriptionAttribute;
                titleKeys.Add(new Tuple<string, string>(prop.Name, description?.Description));
            }
            var titleKeyDictionarys = new Dictionary<string, int>();
            var returnList = new List<T>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts
                                                          .First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>()
                                                             .First();
                var rows = sheetData.Descendants<Row>();
                var cellValues = new List<Dictionary<int, string>>();
                var titleRow = rows.First();
                SharedStringTable strTable = spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable;
                var titleRowValues = titleRow.Descendants<Cell>().Select(c => GetCellText(c, strTable))
                                                                 .ToList();
                foreach (var key in titleKeys)
                {
                    titleKeyDictionarys.Add(key.Item1, titleRowValues.IndexOf(useDescription && !string.IsNullOrEmpty(key.Item2) ? key.Item2 : key.Item1));
                }

                foreach (Row r in rows.Skip(1))
                {
                    if (r.Elements<Cell>()
                         .All(c => string.IsNullOrEmpty(c.CellValue?.InnerText)))
                    {
                        continue;
                    }
                    var cellValue = new Dictionary<int, string>();
                    var cellcount = r.Elements<Cell>()
                                     .Count();
                    for (var i = 0; i < cellcount; i++)
                    {
                        var cell = r.Elements<Cell>()
                                    .ElementAt(i);
                        cellValue.Add(CellReferenceToIndex(cell), GetCellText(cell, strTable));
                    }
                    cellValues.Add(cellValue);
                }

                foreach (Dictionary<int, string> cellValue in cellValues)
                {
                    var data = Activator.CreateInstance(typeof(T));

                    foreach (var property in typeof(T).GetProperties())
                    {
                        if (titleKeyDictionarys.TryGetValue(property.Name, out int index))
                        {
                            if (index == -1)
                                continue;
                            if (cellValue.TryGetValue(index, out var foundValue))
                            {
                                if (property.PropertyType == typeof(int))
                                {
                                    property.SetValue(data, Convert.ToInt32(foundValue));
                                }
                                else if (property.PropertyType == typeof(DateTime))
                                {
                                    property.SetValue(data, DateTime.FromOADate(Convert.ToDouble(foundValue)));
                                }
                                else if (property.PropertyType == typeof(string[]))
                                {
                                    property.SetValue(data, foundValue?.Split(','));
                                }
                                else if (property.PropertyType == typeof(long))
                                {
                                    property.SetValue(data, Convert.ToInt64(foundValue));
                                }
                                else
                                {
                                    property.SetValue(data, foundValue);
                                }
                            }
                        }
                    }
                    returnList.Add((T)data);
                }
            }
            return returnList;
        }

        private string GetCellText(Cell cell, SharedStringTable strTable)
        {
            if (cell.ChildElements.Count == 0)
                return null;
            string val = cell.CellValue.InnerText;
            //若為共享字串時的處理邏輯
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                val = strTable.ChildElements[int.Parse(val)].InnerText;
            return val;
        }

        private static int CellReferenceToIndex(Cell cell)
        {
            int index = 0;
            string reference = cell.CellReference.ToString()
                                                 .ToUpper();
            foreach (char ch in reference)
            {
                if (char.IsLetter(ch))
                {
                    int value = (int)ch - (int)'A';
                    index = (index == 0) ? value : ((index + 1) * 26) + value;
                }
                else
                    return index;
            }
            return index;
        }
    }
}
