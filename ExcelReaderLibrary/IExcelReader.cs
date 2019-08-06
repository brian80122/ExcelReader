using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderLibrary
{
    public interface IExcelReader<T> where T :class
    {
        List<T> ReadExcel(string filePath, bool useDescription);
    }
}
