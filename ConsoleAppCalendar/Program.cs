using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NDesk.Options;

namespace ConsoleAppCalendar
{
    class Program
    {
        //parse command line arguments http://www.ndesk.org/Options
        static void Main(string[] args) {
            System.DateTime firstDate;
            System.DateTime secondDate;
            String templateFilePath;
            String outputPath;
            bool show_help=false;
            OptionSet p = new OptionSet()
               .Add("date1|date1=", "First {DATE1} of span to generate calendar", delegate (string v) {
                    if (v != null) {
                       firstDate = System.DateTime.Parse(v);
                   }
               })
               .Add("date2|date2=", "Second {DATE2} of span to generate calendar", delegate (string v) {
                   if (v != null) {
                       secondDate = System.DateTime.Parse(v);
                   }
               })
               .Add("h|?|help", "Show help message", delegate (string v) { show_help = v != null; })
               .Add("template|template=", "{TEMPLATE} xml file", delegate (string v) {
                   if (v != null) {
                       templateFilePath = v;
                   }
               })
               .Add("output|output=", "{RESULT} file with calendar: docx", delegate (string v) {
                   if (v != null)
                   {
                       outputPath = v;
                   }
               });
            if (show_help) {
                DisplayHelpMessage(p);
            }
            List<string> extra = p.Parse(args);
        }

        private static void DisplayHelpMessage(OptionSet p)
        {
            System.Console.WriteLine("== A calendar app. Construct tabletop calendar ==");
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }
    }
}
