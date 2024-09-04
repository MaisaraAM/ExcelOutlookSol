using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinAppDriver.Utilities
{
    public static class HandleExceptions
    {
        public static void LogAnyExceptionAndFailTestCase(Exception e,string message =null)
        {
            //We need to implement log and call it here
            Assert.Fail( (message ==null ? " " : message ) +  e.StackTrace);
        }

    }
}
