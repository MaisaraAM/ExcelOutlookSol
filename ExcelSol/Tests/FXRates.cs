using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Castle.Components.DictionaryAdapter.Xml;
using ExcelSol.Pages;
using NUnit.Framework;
using DataTable = System.Data.DataTable;

namespace ExcelSol.Tests
{
    public class FXRates : TestFixtureBase
    {
        [Test]
        public void getFXRates()
        {
            string filename;
            string url;

            EmailPage.searchEmailAndDownlaodAttachments("Intraday FX rates report", out filename, out url, true);

            RatesPage ratesPage = new RatesPage();

            string path = ratesPage.getExcelPath("Downloads", "ITD_FX_RATE_20240902.xlsx");
            string newPath = ratesPage.getExcelPath("Downloads", "New FX File.xlsx");

            ratesPage.getFXRate(path, newPath, new List<string> { "'USD/AED'", "'USD/BDT'", "'EUR/USD'", "'USD/ARS'", "'USD/BGN'", "'USD/BHD'" });
        }
    }
}
