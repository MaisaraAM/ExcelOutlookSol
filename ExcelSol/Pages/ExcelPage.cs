using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace ExcelSol.Pages
{
    public class ExcelPage : TestFixtureBase
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

        public void loadExcelSheetDT(string excelFilePath, out DataTable excelDT)
        {
            excelDT = new DataTable();
            excelDT.Columns.Add("S/N", typeof(string));
            excelDT.Columns.Add("Make", typeof (string));
            excelDT.Columns.Add("Model", typeof(string));
            excelDT.Columns.Add("Year", typeof(string));
            excelDT.Columns.Add("Chassis #", typeof(string));
            excelDT.Columns.Add("Motor #", typeof(string));
            excelDT.Columns.Add("License #", typeof(string));
            excelDT.Columns.Add("Sylndr Invoice #", typeof(string));
            excelDT.Columns.Add("Credit Control", typeof(string));
            excelDT.Columns.Add("Legal Dep't.", typeof(string));
            excelDT.Columns.Add("Finance Dep't.", typeof(string));

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
                if (index > 4)
                    excelDT.Rows.Add(dr[1], dr[2], dr[3], dr[4], dr[6], dr[7],dr[8], dr[9], dr[18], dr[19], dr[20]);
                index++;
            }

            dr.Close();
            con.Close();
        }

        public void retrieveCells(DataTable dTable)
        {
            DataRow[] dataRows0 = dTable.Select("Year='2021'");
            string make = dataRows0[3]["Model"].ToString();

            string carMake = dTable.Rows[5]["Make"].ToString();
            string carModel = dTable.Rows[5]["Model"].ToString();
            string chassisNum = dTable.Rows[5]["Chassis #"].ToString();
        }

        public void updateCellValue(string excelFilePath, DataTable dTable)
        {
            excelApi = new ExcelApi(excelFilePath);
            excelApi.OpenExcel();
            List<string> sheetList = excelApi.getSheetName();

            excelApi.UpdateCellData(sheetList[0], 5, 9, "Corolla");
            excelApi.CloseExcel();

            string newModel = dTable.Rows[3]["Year"].ToString();
        }

        public void updateLicense(string excelFilePath, DataTable dTable, string sNum, string licenseNum)
        {
            excelApi = new ExcelApi(excelFilePath);
            excelApi.OpenExcel();
            List<string> sheetList = excelApi.getSheetName();

            int tCount = dTable.Rows.Count;

            for (int i = 0; i <= tCount; i++)
            {
                if (dTable.Rows[i]["S/N"].ToString() == sNum)
                {
                    excelApi.UpdateCellData(sheetList[0], 9, i + 6, licenseNum);
                    break;
                }
            }
            
            excelApi.CloseExcel();

            string lic = dTable.Rows[22]["License #"].ToString();
        }

        public void insertRow(DataTable dTableMain, string excelFilePath)
        {
            excelApi = new ExcelApi(excelFilePath);
            excelApi.OpenExcel();
            List<string> sheetList = excelApi.getSheetName();

            for (int i = 0; i < dTableMain.Rows.Count; i++)
            {
                DataRow dataRow = dTableMain.Rows[i];

                excelApi.UpdateCellData(sheetList[0], 1, i + 1, dataRow["S/N"].ToString());
                excelApi.UpdateCellData(sheetList[0], 2, i + 1, dataRow["Make"].ToString());
                excelApi.UpdateCellData(sheetList[0], 3, i + 1, dataRow["Model"].ToString());
                excelApi.UpdateCellData(sheetList[0], 4, i + 1, dataRow["Year"].ToString());
                excelApi.UpdateCellData(sheetList[0], 5, i + 1, dataRow["Chassis #"].ToString());
                excelApi.UpdateCellData(sheetList[0], 6, i + 1, dataRow["Motor #"].ToString());
                excelApi.UpdateCellData(sheetList[0], 7, i + 1, dataRow["License #"].ToString());
                excelApi.UpdateCellData(sheetList[0], 8, i + 1, dataRow["Sylndr Invoice #"].ToString());
                excelApi.UpdateCellData(sheetList[0], 9, i + 1, dataRow["Credit Control"].ToString());
                excelApi.UpdateCellData(sheetList[0], 10, i + 1, dataRow["Legal Dep't."].ToString());
                excelApi.UpdateCellData(sheetList[0], 11, i + 1, dataRow["Finance Dep't."].ToString());
            }

            excelApi.CloseExcel();
        }
    }
}
