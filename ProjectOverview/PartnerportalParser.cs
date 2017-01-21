using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectOverview
{
   class PartnerportalParser
   {
      private String _startPattern = "no.	!";
      private Dictionary<String, Dictionary<String, Dictionary<String, DateTime>>> m_Data = new Dictionary<string, Dictionary<string, Dictionary<string, DateTime>>>();

      public Dictionary<String, Dictionary<String, Dictionary<String, DateTime>>> ReturnData() { return m_Data; }
      private String GetClipBoardText()
      {
         String idat = null;
         Exception threadEx = null;
         Thread staThread = new Thread(
             delegate()
             {
                try
                {
                   idat = Clipboard.GetText(TextDataFormat.Text);
                }

                catch (Exception ex)
                {
                   threadEx = ex;
                }
             });
         staThread.SetApartmentState(ApartmentState.STA);
         staThread.Start();
         staThread.Join();
         return idat;
      }

      public void ParseClipBoard()
      {
         String ClipboardText = GetClipBoardText();
         String PatternMusterPhase = @"(^[eE]{1}[0-9]{3}[a-d]{1})|(^[eE]{1}[0-9]{3}.[0-9]{1,2}[a-d]{0,1})|(^[eE]{1}[0-9]{3})|(^[Xx]{1}[0-9]{3}.[0-9]{1,2}[a-d]{0,1})|(^[Xx]{1}[0-9]{3})|(^BrabusAMG)|(^Brabus)";
         String PatternRelease = @"(\.[rR]el[0-9]{0,2}$)|([rR]el[0-9]{0,2}$)|([Pp]re[0-9]{2}$)|([Pp]re [0-9]{2}$)";

         StringBuilder sb = new StringBuilder();

         StringWriter sw = new StringWriter(sb);
         sw.Write(ClipboardText);
         sw.Close();

         StringReader reader = new StringReader(sb.ToString());
         while (reader.Peek()> 0)
         {
            if (reader.ReadLine().Contains(_startPattern))
            {
               while (reader.Peek() > 0)
               {
                  string Row = reader.ReadLine();
                  List<string> RowContent = Row.Split('\t').ToList();

                  if (m_Data.ContainsKey(RowContent[1]))
                  {
                     Dictionary<String, Dictionary<String, DateTime>> m;
                     m_Data.TryGetValue(RowContent[1], out m);

                     Regex MusterPhase2 = new Regex(PatternMusterPhase);
                     Regex Release2 = new Regex(PatternRelease);
                     String musterphase2 = MusterPhase2.Match(RowContent[2]).Value;
                     String release2 = Release2.Match(RowContent[2]).Value;


                     if (m.ContainsKey(musterphase2))
                     {
                        Dictionary<String, DateTime> mm;
                        Regex MusterPhase = new Regex(PatternMusterPhase);
                        Regex Release = new Regex(PatternRelease);
                        String musterphase = MusterPhase.Match(RowContent[2]).Value;
                        String release = Release.Match(RowContent[2]).Value;

                        if (release == "")
                           Console.WriteLine(RowContent[2]);
                        else
                        {
                           m.TryGetValue(musterphase, out mm);
                           mm.Add(release, DateTime.Parse(RowContent[3]));
                        }

                     }
                     else 
                     {
                           var data = new Dictionary<String, DateTime>();
                           Regex MusterPhase = new Regex(PatternMusterPhase);
                           Regex Release = new Regex(PatternRelease);
                           String musterphase = MusterPhase.Match(RowContent[2]).Value;
                           String release = Release.Match(RowContent[2]).Value;
                           if (release == "")
                              Console.WriteLine(RowContent[2]);
                           else
                           {
                              data.Add(release, DateTime.Parse(RowContent[3]));
                              m.Add(musterphase, data);
                           }

                     }
                     
                  }
                  else 
                  {
                     var data = new Dictionary<String, DateTime>();

                     Regex MusterPhase = new Regex(PatternMusterPhase);
                     Regex Release = new Regex(PatternRelease);
                     String musterphase = MusterPhase.Match(RowContent[2]).Value;
                     String release = Release.Match(RowContent[2]).Value;

                     if (release == "")
                        Console.WriteLine(RowContent[2]);
                     else
                     {
                        data.Add(release, DateTime.Parse(RowContent[3]));
                        var data2 = new Dictionary<String, Dictionary<String, DateTime>>();
                        data2.Add(musterphase, data);
                        m_Data.Add(RowContent[1], data2);
                     }
                  }

               }
            }
         }
      }

   }
}
