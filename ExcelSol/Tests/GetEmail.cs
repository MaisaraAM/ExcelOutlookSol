using ExcelSol.Pages;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSol.Tests
{
    public class GetEmail : TestFixtureBase
    {
        [Test]
        public void email()
        {
            string attach = "646";
            string url = "";

            //EmailPage.searchEmailAndDownlaodAttachments("RE: Test Attached Subject", out attach, out url, true);
            EmailPage.sendEmailResults("Test send attachment", "t-mamaher@EFG-HERMES.com", "", "Test test", true);
        }
    }
}
