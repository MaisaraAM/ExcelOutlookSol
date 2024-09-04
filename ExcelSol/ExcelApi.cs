using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using xl = Microsoft.Office.Interop.Excel;

namespace ExcelSol
{
   public class ExcelApi : TestFixtureBase
    {

        xl.Application xlApp = null;
        xl.Workbooks workbooks = null;
        xl.Workbook workbook = null;
        Hashtable sheets;
        Hashtable sheets_Hidden;
        Hashtable sheets_All;
        public static bool Skip = false;
        public string xlFilePath;

        public ExcelApi(string xlFilePath)
        {
            this.xlFilePath = xlFilePath;
        }
        public ExcelApi()
        {

        }

        public void OpenExcel()
        {
            xlApp = new xl.Application();
            workbooks = xlApp.Workbooks;
            workbook = workbooks.OpenXML(xlFilePath);
            sheets = new Hashtable();
            sheets_Hidden = new Hashtable();
            sheets_All = new Hashtable();

            int count = 1;
            //Storing worksheet names in Hashtable.
            foreach (xl.Worksheet sheet in workbook.Sheets)
            {

                if (sheet.Visible.ToString() != "xlSheetHidden")
                {
                    sheets[count] = sheet.Name;
                    workbook.Sheets[count].Select();
                    break;

                }
                count++;
            }
            count = 1;
            foreach (xl.Worksheet sheet in workbook.Sheets)
            {

                if (sheet.Visible.ToString() == "xlSheetHidden")
                {
                    sheets_Hidden[count] = sheet.Name;

                }
                else
                    break;

                count++;
            }
            count = 1;
            foreach (xl.Worksheet sheet in workbook.Sheets)
            {
                sheets_All[count] = sheet.Name;

                count++;
            }

        }
        public void Close_excel_process()
        {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("Excel");
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
        }
        public void CloseExcel()
        {
            workbook.Close(false, xlFilePath, null); // Close the connection to workbook
            Marshal.FinalReleaseComObject(workbook); // Release unmanaged object references.
            workbook = null;

            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            workbooks = null;

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            xlApp = null;
            // Close_excel_process();
        }


        public string GetCellData(string sheetName, int colNumber, int rowNumber)
        {
            // OpenExcel();

            string value = string.Empty;
            int sheetValue = 0;

            if (sheets_All.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets_All)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;

                value = Convert.ToString((range.Cells[rowNumber, colNumber] as xl.Range).Text);
                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            //  CloseExcel();
            return value;
        }
        public double GetCellDataDecimalValue(string sheetName, int colNumber, int rowNumber)
        {
            // OpenExcel();

            // string value = string.Empty;
            double value = 0;
            int sheetValue = 0;

            if (sheets_All.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets_All)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;

                try
                {
                    value = Convert.ToDouble((range.Cells[rowNumber, colNumber] as xl.Range).Value2);
                }
                catch { value = 0; }
                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            //  CloseExcel();
            return value;
        }
        public string GetCellDataValue(string sheetName, int colNumber, int rowNumber)
        {
            // OpenExcel();

            string value = string.Empty;
            int sheetValue = 0;

            if (sheets_All.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets_All)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;

                value = Convert.ToString((range.Cells[rowNumber, colNumber] as xl.Range).Value);
                if (String.IsNullOrEmpty(value))
                    value = "0";
               else if (value.Trim()=="-")
                    value = "0";

                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            //  CloseExcel();
            return value;
        }
        public string GetCellFormula(string sheetName, int colNumber, int rowNumber)
        {
            // OpenExcel();

            string value = string.Empty;
            int sheetValue = 0;

            if (sheets_All.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets_All)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;

                value = Convert.ToString((range.Cells[rowNumber, colNumber] as xl.Range).Formula);
                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            //  CloseExcel();
            return value;
        }

        public void UpdateCellData(string sheetName, int colNumber, int rowNumber, string Value)
        {
            //  OpenExcel();
            string value = string.Empty;
            int sheetValue = 0;
            if (sheets_All.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets_All)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;

                range.Cells[rowNumber, colNumber] = Value;
                workbook.Save();
            }
            //  CloseExcel();
        }
        public void UpdateCellFolmula(string sheetName, int colNumber, int rowNumber, string Value)
        {
            //  OpenExcel();
            string value = string.Empty;
            int sheetValue = 0;
            if (sheets_All.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets_All)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;

                range.Cells[rowNumber, colNumber].Formula = Value;
                workbook.Save();
            }
            //  CloseExcel();
        }
        public void UpdateCellData(string sheetName, int colNumber, int rowNumber, string Value, out int rows_count)
        {
            //  OpenExcel();
            string value = string.Empty;
            int sheetValue = 0;
            rows_count = 0;
            if (sheets_All.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets_All)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;
                rows_count = worksheet.UsedRange.Rows.Count;
                range.Cells[rowNumber, colNumber] = Value;
                workbook.Save();
            }
            //  CloseExcel();
        }
        public string Get_excel_Downloads_path(string FileName)
        {
            var application_path = AppDomain.CurrentDomain.SetupInformation.ApplicationBase.Replace("\\bin\\Debug", "\\Downloads");
            return $@"{application_path}{FileName}.xlsx";
        }
        public static string Get_PDF_Downloads_path(string FileName)
        {
            var application_path = AppDomain.CurrentDomain.SetupInformation.ApplicationBase.Replace("\\bin\\Debug", "\\Downloads");
            return $@"{application_path}{FileName}.pdf";

        }

        public string Get_excel_Downloads_path_CR(string FileName)
        {
            var application_path = AppDomain.CurrentDomain.SetupInformation.ApplicationBase.Replace("\\bin\\Debug", "\\Crystal_reports");
            return $@"{application_path}{FileName}.xls";

        }



        public static string Get_testdata_excel_path(string FileName)
        {
            var application_path = AppDomain.CurrentDomain.SetupInformation.ApplicationBase.Replace("\\bin\\Debug", "\\Files");
            return $@"{application_path}{FileName}.xlsx";
        }
        public static string Get_excel_Uploads_path(string FileName)
        {
            var application_path = AppDomain.CurrentDomain.SetupInformation.ApplicationBase.Replace("\\bin\\Debug", "\\Uploads");
            return $@"{application_path}{FileName}.xlsx";

        }


      
        public List<string> getTestDataFromExcelSheetByColumn(string Sheet_name, int columnNumber, int startRowIndex, int rowsCount = 0)
        {
            List<string> SheetName = getSheetName();
            int sheetValue = 0;
            if (rowsCount == 0)
            {
                if (sheets_All.ContainsValue(SheetName[0].ToString()))
                {
                    foreach (DictionaryEntry sheet in sheets_All)
                    {
                        if (sheet.Value.Equals(SheetName[0].ToString()))
                        {
                            sheetValue = (int)sheet.Key;
                        }
                    }
                    xl.Worksheet worksheet = null;
                    worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                    xl.Range range = worksheet.UsedRange;


                    rowsCount = worksheet.UsedRange.Rows.Count;
                }
            }
            List<string> testDataList = new List<string>();
            for (int i = startRowIndex; i < startRowIndex + rowsCount - 1; i++)
            {

                testDataList.Add(GetCellDataValue(Sheet_name, columnNumber, i));
            }
            Assert.IsTrue(testDataList.Count > 0, "Couldn't getTestDataFromExcelSheetByColumn Number: { " + columnNumber + " }");
            return testDataList;
        }

        public List<string> getTestDataFromExcelSheetByRow(string Sheet_name, int rowNumber, int startColIndex, int colsCount)
        {
            List<string> testDataList = new List<string>();
            for (int i = startColIndex; i < startColIndex + colsCount; i++)
            {
                testDataList.Add(GetCellData(Sheet_name, i, rowNumber));
            }
            return testDataList;
        }
        public int GetRowsCountUsedRangePerColumn(string sheetName, int ColRowNumber, bool byCol = true)
        {

            int sheetValue = 0;
            int rows_count = 0;
            if (sheets_All.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets_All)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                //xl.Range range = worksheet.UsedRange;
                if (byCol)
                    rows_count = worksheet.UsedRange[worksheet.Columns[ColRowNumber]].Rows.Count;
                else
                    rows_count = worksheet.UsedRange[worksheet.Rows[ColRowNumber]].Rows.Count;
            }
            return rows_count;
        }


        public static void Delete_file(string root_path)
        {
            try
            {
                if (File.Exists(Path.Combine(root_path)))
                {
                    File.Delete(Path.Combine(root_path));
                }
            }
            catch (IOException Exp)
            {
                Console.WriteLine(Exp.Message);
            }
        }



        //New functions
        //---------------------------
        public List<string> getAllFolderUnderTheRoot(string Path)
        {
            Thread.Sleep(1000);
            List<string> dirs = Directory.GetDirectories(Path, "*", SearchOption.TopDirectoryOnly).ToList();
            return dirs;
        }

        public List<string> getAllExcelFiles(string Path)
        {
            List<string> dirs = getAllFolderUnderTheRoot(Path);

            List<string> AllFiles = new List<string>();
            foreach (string dir in dirs)
            {
                List<string> files = Directory.GetFiles(dir, "*.xlsx").ToList();
                files.RemoveAll(s => s.Contains("~$"));
                AllFiles.AddRange(files);

            }

            return AllFiles;

        }
        public List<string> getAllFiles(string Path)
        {
            Thread.Sleep(1000);
            List<string> AllFiles = new List<string>();
            List<string> files = Directory.GetFiles(Path).ToList();
            files.RemoveAll(s => s.Contains("~$"));
            AllFiles.AddRange(files);

            return AllFiles;

        }
        public List<string> getAllExcelFilesDir(string Path)
        {

            List<string> AllFiles = new List<string>();
            List<string> files = Directory.GetFiles(Path, "*.xlsx").ToList();
            files.RemoveAll(s => s.Contains("~$"));
            AllFiles.AddRange(files);

            return AllFiles;

        }
        public List<string> getAllExcelFilesUnderRootandSubroot(string Path)
        {

            List<string> files = Directory.GetFiles(Path, "*.xlsx", SearchOption.AllDirectories).ToList();

            return files;

        }
     
        public static string getExcelpath(string FolderName, string FileName)
        {

            var application_path = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            var applicationPath_new = application_path.Replace("\\bin\\Debug", "");
            var applicationPath_new_name = applicationPath_new + "\\"+FolderName;
            applicationPath_new_name = @applicationPath_new_name.Replace(@"\\"+FolderName, @"\"+FolderName);
            string final_path = $@"{applicationPath_new_name}\{FileName}"+".xlsx";
            return final_path;
        }
        public static string getExcelFolderpath(string FolderName)
        {
            var application_path = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            var applicationPath_new = application_path.Replace("\\bin\\Debug", "");
            var applicationPath_new_name = applicationPath_new + "\\" + FolderName;
            applicationPath_new_name = @applicationPath_new_name.Replace(@"\\" + FolderName, @"\" + FolderName);
            string final_path = $@"{applicationPath_new_name}";
            return final_path;
        }
        public List<string> getListColumByName(string sheetName, int columnNumber, string txt, out int headerTxtIndex, bool checkRowsLimit = true)
        {
            Skip = false;

            int rowsCount = GetRowsCountUsedRangePerColumn(sheetName);

            if (rowsCount > 500 && checkRowsLimit)
                rowsCount = 250;


            List<string> alldataList = new List<string>();
            List<DateTime> alldataListdt = new List<DateTime>();
            List<string> dataList = new List<string>();
            headerTxtIndex = 1;
            bool headerFound = false;

            for (int i = 1; i <= rowsCount; i++)
                alldataList.Add(GetCellData(sheetName, columnNumber, i));

            for (int i = 0; i < alldataList.Count; i++)
            {
                if (!(alldataList[i] is null))
                {
                    if (alldataList[i].ToString().Trim().Equals(txt, StringComparison.OrdinalIgnoreCase))
                    {
                        headerFound = true;
                        headerTxtIndex = i + 1;
                        break;
                    }
                }
            }

            for (int i = headerTxtIndex; i < alldataList.Count; i++)
            {
                if (alldataList[i].Trim() == "")
                    break;

                dataList.Add(alldataList[i].ToString().Trim());

            }
            if (dataList.Count == 0 || headerFound == false)
            {
                Skip = true;

            }
            return dataList;
        }



        public List<double> getListColumByNameDecimalValues(string sheetName, int columnNumber, string txt, out int headerTxtIndex, bool checkRowsLimit = true)
        {
            Skip = false;

            int rowsCount = GetRowsCountUsedRangePerColumn(sheetName);

            if (rowsCount > 500 && checkRowsLimit)
                rowsCount = 250;


            List<double> alldataList = new List<double>();
            List<DateTime> alldataListdt = new List<DateTime>();
            List<double> dataList = new List<double>();
            headerTxtIndex = 1;
            bool headerFound = false;

            for (int i = 1; i <= rowsCount; i++)
                alldataList.Add(GetCellDataDecimalValue(sheetName, columnNumber, i));

            for (int i = 0; i < alldataList.Count; i++)
            {
                if (!(alldataList[i].ToString() is null))
                {
                    if (alldataList[i].ToString().Trim().Equals(txt, StringComparison.OrdinalIgnoreCase))
                    {
                        headerFound = true;
                        headerTxtIndex = i + 1;
                        break;
                    }
                }
            }

            for (int i = headerTxtIndex; i < alldataList.Count; i++)
            {
                if (alldataList[i].ToString().Trim() == "")
                    break;

                dataList.Add(alldataList[i]);

            }
            if (dataList.Count == 0 || headerFound == false)
            {
                Skip = true;

            }
            return dataList;
        }


        public List<string> getSheetName()
        {
            List<string> sheetName = new List<string>();

            foreach (xl.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Visible.ToString() != "xlSheetHidden")
                {
                    sheetName.Add(sheet.Name);
                }
            }

            return sheetName;
        }

        //-------------------------
        public int GetRowsCountUsedRangePerColumn(string sheetName)
        {

            int sheetValue = 0;
            int rows_count = 0;
            if (sheets_All.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets_All)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                
                xl.Worksheet worksheet = null;               
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
               
                xl.Range range = worksheet.UsedRange;
                worksheet.AutoFilterMode = false;
                rows_count = worksheet.UsedRange.Rows.Count;
            }
            return rows_count;
        }
        public List<string> getTestDataFromExcelSheetByColumnUsedRange(string Sheet_name, int columnNumber, int startRowIndex = 2)
        {
            int rowsCount = GetRowsCountUsedRangePerColumn(Sheet_name);
            List<string> testDataList = new List<string>();

            for (int i = startRowIndex; i <= rowsCount; i++)
            {
                testDataList.Add(GetCellData(Sheet_name, columnNumber, i));
            }

            return testDataList;
        }
        public List<string> getDataFromExcelSheetColumnByRows(string Sheet_name, int columnNumber, int startRowNumber ,int endRowNumber)
        {        
            List<string> DataList = new List<string>();
            for (int i = startRowNumber; i <= endRowNumber; i++)
            {
                DataList.Add(GetCellData(Sheet_name, columnNumber, i));
            }
            return DataList;
        }
        public List<double> getDataFromExcelSheetColumnByRowsDecimalValues(string Sheet_name, int columnNumber, int startRowNumber, int endRowNumber)
        {
            List<double> DataList = new List<double>();
            for (int i = startRowNumber; i <= endRowNumber; i++)
            {
                DataList.Add(GetCellDataDecimalValue(Sheet_name, columnNumber, i));
            }
            return DataList;
        }

        public List<string> getDataFromExcelSheetColumnByRowsFormula(string Sheet_name, int columnNumber, int startRowNumber, int endRowNumber)
        {
            List<string> DataList = new List<string>();
            for (int i = startRowNumber; i <= endRowNumber; i++)
            {
                DataList.Add(GetCellFormula(Sheet_name, columnNumber, i));
            }
            return DataList;
        }

        public void addRowInExcel(List<int> colNumber, List<string> values)
        {
            List<string> SheetName = getSheetName();
            int sheetValue = 0;
            int rows_count = 0;
            if (sheets_All.ContainsValue(SheetName[0].ToString()))
            {
                foreach (DictionaryEntry sheet in sheets_All)
                {
                    if (sheet.Value.Equals(SheetName[0].ToString()))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;

                rows_count = worksheet.UsedRange.Rows.Count + 1;

                for (int i = 0; i < colNumber.Count; i++)
                {
                    range.Cells[rows_count, colNumber[i]] = values[i];
                    range.Cells[rows_count, 8].Formula = "=F" + rows_count.ToString() + "-G" + rows_count.ToString();
                    range.Cells[rows_count, 11].Formula = "=I" + rows_count.ToString() + "-J" + rows_count.ToString();
                }
                workbook.Save();
            }
        }
        public void addRowInExcelWithoutFormula(List<int> colNumber, List<string> values,int sheetNumber=0)
        {
            List<string> SheetName = getSheetName();
            int sheetValue = 0;
            int rows_count = 0;
            if (sheets_All.ContainsValue(SheetName[sheetNumber].ToString()))
            {
                foreach (DictionaryEntry sheet in sheets_All)
                {
                    if (sheet.Value.Equals(SheetName[sheetNumber].ToString()))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;

                rows_count = worksheet.UsedRange.Rows.Count + 1;

                for (int i = 0; i < colNumber.Count; i++)
                {
                    range.Cells[rows_count, colNumber[i]] = values[i];
                }
                workbook.Save();
            }
        }
        public void deleteRowsInExcel(int numSkipRows = 1)
        {
            List<string> SheetName = getSheetName();
            int sheetValue = 0;
            int rows_count = 0;
            for (int k = 0; k < SheetName.Count; k++)
            {
                if (sheets_All.ContainsValue(SheetName[k].ToString()))
                {
                    foreach (DictionaryEntry sheet in sheets_All)
                    {
                        if (sheet.Value.Equals(SheetName[k].ToString()))
                        {
                            sheetValue = (int)sheet.Key;
                        }
                    }
                    xl.Worksheet worksheet = null;
                    worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                    xl.Range range = worksheet.UsedRange;


                    rows_count = worksheet.UsedRange.Rows.Count;
                    for (int i = numSkipRows + 1; i <= rows_count; i++)
                    {
                        worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                        range = worksheet.UsedRange;
                        range.Cells[2, 1].EntireRow.Delete();
                    }
                    workbook.Save();
                }
            }
        }
        public Worksheet getWorkSheet(string sheetName)
        {
            xl.Worksheet worksheet = null;
            int index = getSheetName().IndexOf(sheetName) + 1;
            worksheet = workbook.Sheets[index + sheets_Hidden.Count] as xl.Worksheet;

            return worksheet;
        }
        public int getsheetindex()
        {
            int count = 1;
            int index = 0;
            // Storing worksheet names in Hashtable.
            foreach (xl.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Visible.ToString() != "xlSheetHidden")
                {
                    index = workbook.Sheets[count].Index;
                    break;
                }
                count++;
            }
            return index;

        }


        public static string checkColumnNameHaveEmpytCells(int expectedCount, IList<int> expectedListCount, IList<string> headerExcelNameList)
        {
            Assert.AreEqual(expectedListCount.Count, headerExcelNameList.Count, "The Count of expected List is not equal count of header Excel Name list");
            string columnNameHaveEmpytCells = "";

            for (int i = 0; i < expectedListCount.Count; i++)
            {
                if (expectedListCount[i] != expectedCount)
                {
                    columnNameHaveEmpytCells += headerExcelNameList[i] + " , ";
                }
            }

            return columnNameHaveEmpytCells;
        }

    }
}



