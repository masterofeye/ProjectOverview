using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ProjectOverview
{
   class Program
   {
      static void Main(string[] args)
      {
         Dictionary<String, String> test = new Dictionary<string, string>();

         PartnerportalParser p = new PartnerportalParser();
         p.ParseClipBoard();

         ExcelGenerator excel = new ExcelGenerator(p.ReturnData());

      }
   }
}
