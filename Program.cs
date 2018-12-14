using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


namespace vlookup
{
    class Program
    {
        static void Main(string[] args)
        {


            //sorting(@"C:\temp\test.xlsx","Sayfa1","B1","B100000");
            //autofilter(@"C:\temp\test.xlsx","Sayfa1","A1:D1",1,55,5);

        }
        public static void sorting(string excel, string sheetname,string sortrange1, string sortrange2)
        {
            //excel = @"C:\temp\test.xlsx";
            var exApp = new Microsoft.Office.Interop.Excel.Application();
            exApp.Visible = true;


            var exWbk = exApp.Workbooks.Open(excel);

            var exWks = (Microsoft.Office.Interop.Excel.Worksheet)exWbk.Sheets[sheetname];


            exWks.Sort.SetRange(exWks.Range["A1", "G100000"]);
            exWks.Sort.Header = XlYesNoGuess.xlYes;
            exWks.Sort.SortFields.Add(exWks.Range[sortrange1, sortrange2], XlSortOn.xlSortOnValues, XlSortOrder.xlDescending);


            exWks.Sort.Apply();


            exWbk.Save();
          

            exApp.Workbooks.Close();
            
            exApp.Quit();
         
        }
        public static void vlookup(string contractDetailReport, string envanterlist, string sourceexcelsheet, string destexcelsheet)
        {
            var spreadsheetLocation = contractDetailReport;
            var spreadsheetLocationSource = envanterlist;


            var exApp = new Microsoft.Office.Interop.Excel.Application();
            exApp.Visible = true;

            var exWbk = exApp.Workbooks.Open(spreadsheetLocation);

            var exWks = (Microsoft.Office.Interop.Excel.Worksheet)exWbk.Sheets[sourceexcelsheet];


            var exAppSource = new Microsoft.Office.Interop.Excel.Application();
            exAppSource.Visible = true;

            var exWbkSource = exAppSource.Workbooks.Open(spreadsheetLocationSource);

            var exWksSource = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.Sheets[destexcelsheet];


            //VLOOKUP A HAZIRLIK

            Range myrange = exWbk.Sheets[sourceexcelsheet].Range("b2:o300000");
            Range myrange1 = exWbkSource.Sheets[destexcelsheet].Range("A4:A145");

            var returnValue = exApp.WorksheetFunction.VLookup(myrange1, myrange, 14, false);

            Worksheet oSheet = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.ActiveSheet;
            oSheet.get_Range("B4", "B145").Value2 = returnValue;


            Range returninsert = exWbkSource.Sheets[destexcelsheet].Range("B4:B145");
            Range myrangeafterinsert = exWbk.Sheets[sourceexcelsheet].Range("o2:p300000");

            var returnValueafterInsert = exApp.WorksheetFunction.VLookup(returninsert, myrangeafterinsert, 2, false);


            Worksheet oSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.ActiveSheet;
            oSheet2.get_Range("C4", "C145").Value2 = returnValueafterInsert;


            Range myrange3 = exWbk.Sheets[sourceexcelsheet].Range("o2:y300000");
            var returnvalue3 = exApp.WorksheetFunction.VLookup(returninsert, myrange3, 11, false);

            Worksheet oSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.ActiveSheet;
            oSheet3.get_Range("D4", "D145").Value2 = returnvalue3;

            Range myrange4 = exWbk.Sheets[sourceexcelsheet].Range("o2:AC300000");
            var returnvalue4 = exApp.WorksheetFunction.VLookup(returninsert, myrange4, 15, false);

            Worksheet oSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.ActiveSheet;
            oSheet4.get_Range("E4", "E145").Value2 = returnvalue4;

            Range myrange5 = exWbk.Sheets[sourceexcelsheet].Range("o2:AD300000");
            var returnvalue5 = exApp.WorksheetFunction.VLookup(returninsert, myrange5, 16, false);

            Worksheet oSheet5 = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.ActiveSheet;
            oSheet5.get_Range("F4", "F145").Value2 = returnvalue5;

            exWbk.Save();
            exWbkSource.Save();

            exApp.Workbooks.Close();
            exAppSource.Workbooks.Close();

            exApp.Quit();
            exAppSource.Quit();



        }
        public static void autofilter(string excelpath, string sheetname, string range, int positionofcolumn, object value1, object value2 = null)
        {

            var exApp = new Microsoft.Office.Interop.Excel.Application();
            exApp.Visible = true;

            var exWbk = exApp.Workbooks.Open(excelpath);

            var exWks = (Microsoft.Office.Interop.Excel.Worksheet)exWbk.Sheets[sheetname];

            //Worksheet oSheet5 = (Microsoft.Office.Interop.Excel.Worksheet)exWbk.ActiveSheet;
            exWbk.Sheets[sheetname].AutoFilterMode = false;

            exWbk.Sheets[sheetname].Range(range).AutoFilter(positionofcolumn, value1, XlAutoFilterOperator.xlOr, value2, true);

            exWbk.Save();


            exApp.Workbooks.Close();

            exApp.Quit();

        }
    }
}
