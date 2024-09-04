using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


namespace WinAppDriver.Utilities
{
    public class WindowsOSActions
    {
        private WindowsDriver<WindowsElement> _session;
        

        public WindowsOSActions(WindowsDriver<WindowsElement> session)
        {
            _session = session;
        }

        public static void killProcess(string processName)
        {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName(processName);
            foreach (System.Diagnostics.Process p in processes)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
            Thread.Sleep(2000);


        }
        public static void killProcess_New(string processName , DataView DV)
        {

            int processId = 0;

            foreach (DataRowView dr in DV)
            {
                if (dr["ProcessName"].ToString()== processName)
                {
                    processId=int.Parse(dr["ProcessID"].ToString());
                }
                                      
            }

            if (processId != 0)
            {
                try
                {
                    Process processes = Process.GetProcessById(processId);
                    processes.Kill();
                }
                catch { }
            }
           
        }
        public static string GetProcess(string processName, string screenName = "")
        {
           // DriverFactory.Run_WinappDriver();
            IntPtr appTopLevelWindowHandle = new IntPtr();

            foreach (Process clsProcess in Process.GetProcesses())
            {
                
                if ((clsProcess.ProcessName.IndexOf(processName, StringComparison.OrdinalIgnoreCase) >= 0
                  && screenName.Trim().Length == 0) ||
                    (clsProcess.MainWindowTitle.IndexOf(screenName, StringComparison.OrdinalIgnoreCase) >= 0
                  && clsProcess.ProcessName.IndexOf(processName, StringComparison.OrdinalIgnoreCase) >= 0
                  && screenName.Trim().Length > 0))
                {
                   
                    if (clsProcess.MainWindowHandle.ToInt32() == 0)
                        continue;
                    appTopLevelWindowHandle = clsProcess.MainWindowHandle;
                    
                     break;
                }
            }
            return appTopLevelWindowHandle.ToString("x");

        }

        public static List<string> GetProcessList(string processName, string screenName = "", int countOfApps = 2)
        {
            List<string> appTopLevelWindowHandleList = new List<string>();
           // DriverFactory.Run_WinappDriver();
            int count = 0;

            IntPtr appTopLevelWindowHandle = new IntPtr();
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if ((clsProcess.ProcessName.IndexOf(processName, StringComparison.OrdinalIgnoreCase) >= 0
                  && screenName.Trim().Length == 0) ||
                    (clsProcess.MainWindowTitle.IndexOf(screenName, StringComparison.OrdinalIgnoreCase) >= 0
                  && clsProcess.ProcessName.IndexOf(processName, StringComparison.OrdinalIgnoreCase) >= 0
                  && screenName.Trim().Length > 0))
                {
                    count += 1;
                    appTopLevelWindowHandle = clsProcess.MainWindowHandle;
                    appTopLevelWindowHandleList.Add(appTopLevelWindowHandle.ToString("x"));
                    if (count == countOfApps)
                        break;
                }
            }
            return appTopLevelWindowHandleList;
        }



        public static string GetProcess_DB(string processName,DataView DV ,string screenName = "")
        {
           // DriverFactory.Run_WinappDriver();
           

            IntPtr appTopLevelWindowHandle = new IntPtr();

            foreach (DataRowView dr in DV)
            {
                if ((dr["ProcessName"].ToString().IndexOf(processName, StringComparison.OrdinalIgnoreCase) >= 0
                  && screenName.Trim().Length == 0) ||
                    (dr["MainWindowTitle"].ToString().IndexOf(screenName, StringComparison.OrdinalIgnoreCase) >= 0
                  && dr["ProcessName"].ToString().IndexOf(processName, StringComparison.OrdinalIgnoreCase) >= 0
                  && screenName.Trim().Length > 0))
                {
                    if (Convert.ToInt32(dr["MainWindowHandle"].ToString(), 16) == 0)
                        continue;

                    appTopLevelWindowHandle = new IntPtr(Convert.ToInt32(dr["MainWindowHandle"].ToString(), 16)); ;


                    break;
                }
            }

            return appTopLevelWindowHandle.ToString("x");

        }
    }



   













}

