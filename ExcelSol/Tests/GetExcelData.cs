using ExcelSol.Pages;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataTable = System.Data.DataTable;

namespace ExcelSol.Tests
{
    public class GetExcelData : TestFixtureBase
    {
        [Test]
        public void getExlData()
        {
            ExcelPage excelPage = new ExcelPage();

            string path = excelPage.getExcelPath("Files", "DD24 1.xlsx");
            excelPage.loadExcelSheetDT(path, out DataTable excelDT);

            excelPage.retrieveCells(excelDT);

            excelPage.updateCellValue(path, excelDT);
            excelPage.updateLicense(path, excelDT, "8", "wjoi 9617");

            string path2 = excelPage.getExcelPath("Files", "New.xlsx");
            excelPage.loadExcelSheetDT(path2, out DataTable excelDT2);

            excelPage.insertRow(excelDT, path2);
        }
    }
}
