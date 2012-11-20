using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel; 


namespace AutomateBPReporting
{
    class ExcelHelper
    {
        Excel.Worksheet babylon;
        Excel.Worksheet ct;
        Excel.Worksheet brooklyn;
        Excel.Worksheet nj;
        Excel.Application excelApp;
        Excel.Workbook excelWorkbook;
        DateTime reportDate;

        public ExcelHelper()
        {
            this.reportDate = DateTime.Now;
            string workbookPath = "d:/Reports/template1.xls";
            excelApp = new Excel.Application();
            excelApp.Visible = false;
            excelWorkbook = excelApp.Workbooks.Open(workbookPath,
                      0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                      true, false, 0, true, false, false);

            babylon = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item("Babylon");
            //brooklyn = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);
            //ct = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(2);
            //nj = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(3);
        }

        public string SaveAs()
        {
            string path = @"d:\Reports\report-" + reportDate.Month.ToString() + "-" + reportDate.Day.ToString() + ".xls";

            excelWorkbook.SaveAs(path);
            excelWorkbook.Close();
            //excelApp.SaveWorkspace(path);

            return path;
        }

        public void WriteWorkSheet(List<List<object>> data, DateTime reportDate, string terminal)
        {
            Excel.Worksheet sheet = (Excel.Worksheet) excelWorkbook.Worksheets.get_Item(terminal); 
            // stamp report date
            sheet.Cells[2, 6] = reportDate.ToShortDateString();
            this.reportDate = reportDate;

            int startRow = 5;
            foreach (List<object> list in data)
            {
                int startCol = 2;

                foreach (object o in list)
                {
                    sheet.Cells[startRow, startCol] = o;
                    startCol++;
                }
                startRow++;
            }
            
        }

        
      
    }
}
