using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ProjectOverview
{
   class ExcelGenerator
   {
      private static Excel.Workbook _book = null;
      private static Excel.Application _app = null;
      private static Excel.Worksheet _sheet = null;

      private DateTime _now;

      const int StartRow = 4;
      const int StartCol = 3;
      const int MonthOffset = 1;
      const int DayOffset = 2;

      const int ProjectOffset = 4;

      private Dictionary<String, Dictionary<String, Dictionary<String, DateTime>>> m_Data;

      public ExcelGenerator(Dictionary<String, Dictionary<String, Dictionary<String, DateTime>>> Data)
      {
         m_Data = Data;
         _app = new Excel.Application();
         _app.Visible = true;
         _book = _app.Workbooks.Open(@"C:\ProjectOverview.xlsx",Type.Missing,false);
         _sheet = _book.ActiveSheet;

         CreateYear(2017);
         GenerateData();


         Excel.Range c1 = _sheet.Cells[1, DateTime.Now.DayOfYear+StartCol];
         Excel.Range c2 = _sheet.Cells[500, DateTime.Now.DayOfYear + StartCol];
         Excel.Range range = _sheet.get_Range(c1, c2);
         range.Interior.ColorIndex = 3;


         _book.Save();
         _book.Close();
         _app.Quit();
         System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);
      }


      private void CreateYear(Int16 Year)
      { 
         uint MaxMonth = 12;
         int DaysInYear = GetDaysInYear(Year);

         Excel.Range c1 = _sheet.Cells[StartRow, StartCol];
         Excel.Range c2 = _sheet.Cells[StartRow, DaysInYear + StartCol-1];
         Excel.Range range = _sheet.get_Range(c1,c2 );
         range.Merge(true);
         range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
         range.Value = Year.ToString();
         range.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

         for (Int16 i = 1; i <= MaxMonth; i++)
         {
            CreateMonth(Year, i);

         }
         range.Columns.AutoFit();
      }

      private void CreateMonth(Int16 Year, Int16 Month)
      {
         int DaysInMonth = DateTime.DaysInMonth(Year, Month);
         string MonthString = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Month);
         DateTime end = new DateTime(Year, Month, 1);
         Excel.Range c1 = _sheet.Cells[StartRow + MonthOffset, StartCol + end.DayOfYear-1];
         Excel.Range c2 = _sheet.Cells[StartRow + MonthOffset, DaysInMonth + end.DayOfYear - 1 + StartCol - 1];
         Excel.Range range = _sheet.get_Range(c1, c2);
         range.Merge(true);
         range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
         range.Value = MonthString;
         range.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

         //formatieren
         for(Int16 i = 1; i <= DaysInMonth; i++)
         {
            _now = new DateTime(Year, Month, i);
            CreateDay(i);
         }

      }

      private void CreateDay(Int16 Day)
      {

         Excel.Range c1 = _sheet.Cells[StartRow + DayOffset, StartCol + _now.DayOfYear-1];
         c1.Value = Day.ToString();
         c1.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
         c1.Columns.AutoFit();
         //Formatieren
      }


      public void GenerateData()
      {
         var counter = 0;
         foreach (var k in m_Data)
         {
            Dictionary<String, Dictionary<String, DateTime>> lvl2 = k.Value;
            CreateProjekt(k.Key, lvl2, counter);
            counter += 1 + lvl2.Count();
         }
      }

      private void CreateProjekt(String Project, Dictionary<String, Dictionary<String, DateTime>> Data, int Counter)
      {
         Excel.Range c1 = _sheet.Cells[StartRow + ProjectOffset + Counter, StartCol - 2];
         c1.Value = Project;
         c1.Columns.AutoFit();
         var counter2 = Counter;
         foreach(var k in Data)
         {
            Dictionary<String, DateTime> lvl3 = k.Value;
            CreateMusterPhase(k.Key, lvl3, counter2);
            counter2++;
         }

      }

      private void CreateMusterPhase(String Musterphase, Dictionary<String, DateTime> Data, int Counter)
      {
         Excel.Range c1 = _sheet.Cells[StartRow + ProjectOffset + Counter, StartCol - 1];
         c1.Value = Musterphase;

         for (int i = 0; i < Data.Count()-1; i++)
         {
            KeyValuePair<String, DateTime> temp = Data.ElementAt(i);
            KeyValuePair<String, DateTime> beforetemp = Data.ElementAt(i + 1);
            if (temp.Value.Year == 2017 && beforetemp.Value.Year == 2017)
            {
               DateTime startDate = temp.Value;
               DateTime endDate = beforetemp.Value.AddDays(1);
               CreateReleases(temp.Key, startDate, endDate, StartRow + ProjectOffset + Counter);
            }
         }
         KeyValuePair<String, DateTime> temp2 = Data.ElementAt(Data.Count()-1);
         if (temp2.Value.Year == 2017)
         {
            Excel.Range c2 = _sheet.Cells[StartRow + ProjectOffset + Counter, temp2.Value.DayOfYear + StartCol - 1];
            c2.Value = temp2.Key;
            c2.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
            c2.Interior.ColorIndex = 36;
            c2.Columns.AutoFit();
         }

      }

      private void CreateReleases(String Name, DateTime StartDate, DateTime EndDate,int Row)
      {


         Excel.Range c1 = _sheet.Cells[Row, StartDate.DayOfYear +StartCol -1];
         Excel.Range c2 = _sheet.Cells[Row, EndDate.DayOfYear + StartCol - 1];
         Excel.Range range = _sheet.get_Range(c1, c2);
         range.Merge(true);
         range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
         range.Value = Name;
         range.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
         range.Interior.ColorIndex = 45;
         range.Columns.AutoFit();
      }

      public static int GetDaysInYear(int year)
      {
         var thisYear = new DateTime(year, 1, 1);
         var nextYear = new DateTime(year + 1, 1, 1);

         return (nextYear - thisYear).Days;
      }


   }
}
