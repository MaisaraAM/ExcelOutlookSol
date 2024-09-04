using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using DataTable = System.Data.DataTable;
using System.Data;

namespace ExcelSol.Pages
{
    public class RatesPage : TestFixtureBase
    {
        ExcelApi excelApi;

        public string getExcelPath(string FolderName, string FileName)
        {
            var application_path = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            var applicationPath_new = application_path.Replace("\\bin\\Debug", "");
            var applicationPath_new_name = applicationPath_new + "\\" + FolderName;
            applicationPath_new_name = @applicationPath_new_name.Replace(@"\\" + FolderName, @"\" + FolderName);
            string final_path = $@"{applicationPath_new_name}\{FileName}";
            return final_path;
        }

        public void getFXRate(string excelFilePath, string newExcelFilePath, List<string> currList)
        {
            DataTable excelDT = new DataTable();
            excelDT.Columns.Add("Quotation", typeof(string));
            excelDT.Columns.Add("Bid", typeof(string));
            excelDT.Columns.Add("Ask", typeof(string));
            excelDT.Columns.Add("Mid", typeof(string));

            excelApi = new ExcelApi(excelFilePath);
            excelApi.OpenExcel();
            List<string> sheetList = excelApi.getSheetName();
            excelApi.CloseExcel();

            string excelDataQuery = "select * from [" + sheetList[0] + "$]";
            string excelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=No;IMEX=1\";", excelFilePath);

            OleDbConnection con = new OleDbConnection(excelConnectionString);
            OleDbCommand cmd = new OleDbCommand(excelDataQuery, con);
            con.Open();
            OleDbDataReader dr = cmd.ExecuteReader();

            int index = 0;
            while (dr.Read())
            {
                if (index > 0)
                    excelDT.Rows.Add(dr[0], dr[1], dr[2], dr[3]);
                index++;
            }

            dr.Close();
            con.Close();

            excelApi = new ExcelApi(newExcelFilePath);
            excelApi.OpenExcel();
            List<string> sheetList2 = excelApi.getSheetName();
            excelApi.deleteRowsInExcel(0);
            excelApi.CloseExcel();

            string s = String.Join(",", currList);
            DataRow[] dataRow = excelDT.Select("Quotation in (" + s + ")");

            for (int i = 0; i < dataRow.Length; i++)
            {
                excelApi.OpenExcel();
                string p = dataRow[i]["Quotation"].ToString();
                excelApi.UpdateCellData(sheetList2[0], 1, i + 1, dataRow[i]["Quotation"].ToString());
                excelApi.UpdateCellData(sheetList2[0], 2, i + 1, dataRow[i]["Bid"].ToString());
                excelApi.UpdateCellData(sheetList2[0], 3, i + 1, dataRow[i]["Ask"].ToString());
                excelApi.UpdateCellData(sheetList2[0], 4, i + 1, dataRow[i]["Mid"].ToString());
                excelApi.CloseExcel();
            }
            //excelApi.CloseExcel();
        }

        //public void retrieveCells(string curr, DataTable dTable, out DataTable cDT)
        //{
        //    List<string> currList = curr.Split(',').ToList();

        //    cDT = new DataTable();
        //    cDT.Columns.Add("Quotation", typeof(string));
        //    cDT.Columns.Add("Bid", typeof(string));
        //    cDT.Columns.Add("Ask", typeof(string));
        //    cDT.Columns.Add("Mid", typeof(string));

        //    foreach (string s in currList)
        //    {
        //        foreach (DataRow row in dTable.Select("Quotation like '%" + s + "%'"))
        //        {
        //            cDT.ImportRow(row);
        //        }
        //    }
        //}

        //public void insertRow(DataRow dataRow, List<string> currList, string excelFilePath)
        //{
        //    excelApi = new ExcelApi(excelFilePath);
        //    excelApi.OpenExcel();
        //    excelApi.deleteRowsInExcel(0);

        //    List<string> sheetList = excelApi.getSheetName();
            
        //    for (int i = 0; i < currList.Count; i++)
        //    {
        //        dataRow = currList.Rows[i];

        //        excelApi.UpdateCellData(sheetList[0], 1, i + 1, dataRow["Quotation"].ToString());
        //        excelApi.UpdateCellData(sheetList[0], 2, i + 1, dataRow["Bid"].ToString());
        //        excelApi.UpdateCellData(sheetList[0], 3, i + 1, dataRow["Ask"].ToString());
        //        excelApi.UpdateCellData(sheetList[0], 4, i + 1, dataRow["Mid"].ToString());
        //    }

        //    excelApi.CloseExcel();
        //}
    }
}
