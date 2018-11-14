using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel;

namespace Assignment
{
    public class ExcelReader
    {
        public IWebDriver driver;
        public string ExcelDataReader(int x, int y)
        {

            excel.Application Xapp = new excel.Application();
            excel.Workbook Xworkbook = Xapp.Workbooks.Open(@"C:\Users\BShah\Desktop\Office_Document\SeleniumScript\ExcelReader\Assignment\Assignment\TestData.xlsx");
            excel._Worksheet Xworksheet = Xworkbook.Sheets[1];
            excel.Range Xrange = Xworksheet.UsedRange;
            return Xrange.Cells[x][y].value2;
        }

    }
}
