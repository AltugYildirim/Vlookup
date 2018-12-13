using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;


namespace vlookup
{
    class Program
    {
        static void Main(string[] args)
        {

            var spreadsheetLocation = @"C:\temp\Copy of 9.2_Contract Detail Report.xlsx";
            var spreadsheetLocationSource = @"C:\temp\VodafoneAA\Book2.xlsb";
            //string txtname = @"C:\temp\report.txt";

            var exApp = new Microsoft.Office.Interop.Excel.Application();
            exApp.Visible = true;

            var exWbk = exApp.Workbooks.Open(spreadsheetLocation);

            var exWks = (Microsoft.Office.Interop.Excel.Worksheet)exWbk.Sheets["Sayfa1"];

            //var exWks = (Microsoft.Office.Interop.Excel.Worksheet)exWbk.Sheets["Sayfa1"];

            var exAppSource = new Microsoft.Office.Interop.Excel.Application();
            exAppSource.Visible = true;

            var exWbkSource = exAppSource.Workbooks.Open(spreadsheetLocationSource);

            var exWksSource = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.Sheets["Envanter List"];


            //VLOOKUP A HAZIRLIK

            Range myrange = exWbk.Sheets["Sayfa1"].Range("b2:o300000");
            Range myrange1 = exWbkSource.Sheets["Envanter List"].Range("A4:A146");

            var returnValue = exApp.WorksheetFunction.VLookup(myrange1, myrange, 14, false);

            Worksheet oSheet = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.ActiveSheet;
            oSheet.get_Range("B4", "B146").Value2 = returnValue;


            Range returninsert = exWbkSource.Sheets["Envanter List"].Range("B4:B146");
            Range myrangeafterinsert = exWbk.Sheets["Sayfa1"].Range("o2:p300000");

            var returnValueafterInsert = exApp.WorksheetFunction.VLookup(returninsert, myrangeafterinsert, 2, false);


            Worksheet oSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.ActiveSheet;
            oSheet2.get_Range("C4", "C146").Value2 = returnValueafterInsert;


            //İşlem bittikten sonra txt dosyası oluştur.
            //using (StreamWriter sw = new StreamWriter(txtname))
            //{

            //     sw.WriteLine("Bitti");

            //}


            Range myrange3 = exWbk.Sheets["Sayfa1"].Range("o2:y300000");
            var returnvalue3 = exApp.WorksheetFunction.VLookup(returninsert, myrange3, 11, false);

            Worksheet oSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.ActiveSheet;
            oSheet3.get_Range("D4", "D146").Value2 = returnvalue3;

            Range myrange4 = exWbk.Sheets["Sayfa1"].Range("o2:AC300000");
            var returnvalue4 = exApp.WorksheetFunction.VLookup(returninsert, myrange4, 15, false);

            Worksheet oSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.ActiveSheet;
            oSheet4.get_Range("E4", "E146").Value2 = returnvalue4;

            Range myrange5 = exWbk.Sheets["Sayfa1"].Range("o2:AD300000");
            var returnvalue5 = exApp.WorksheetFunction.VLookup(returninsert, myrange5, 16, false);

            Worksheet oSheet5 = (Microsoft.Office.Interop.Excel.Worksheet)exWbkSource.ActiveSheet;
            oSheet5.get_Range("F4", "F146").Value2 = returnvalue5;


        }
    }
}
