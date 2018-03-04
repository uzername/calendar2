using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NDesk.Options;

using ConsoleAppCalendar.Interface;


namespace ConsoleAppCalendar
{
    class Program
    {
        //parse command line arguments http://www.ndesk.org/Options
        static void Main(string[] args) {
            System.DateTime? firstDate = null;
            System.DateTime? secondDate = null;
            String templateFilePath;
            String outputPath="";

            StatusStructure messageHandlingStatus = new StatusStructure { show_help = false, templateExpected = false, resultUnspecified = false };

            OptionSet p = new OptionSet()
               .Add("d1|date1=", "First {DATE1} of span to generate calendar", delegate (string v) {
                    if (v != null) {
                       firstDate = System.DateTime.Parse(v);
                   }
               })
               .Add("d2|date2=", "Second {DATE2} of span to generate calendar", delegate (string v) {
                   if (v != null) {
                       secondDate = System.DateTime.Parse(v);
                   }
               })
               .Add("h|?|help", "Show help message", delegate (string v) { messageHandlingStatus.show_help = (v != null); })
               .Add("t|template=", "{TEMPLATE} xml file", delegate (string v) {
                   if (v != null) {
                       templateFilePath = v;
                   }
                   else { messageHandlingStatus.templateExpected = true; }
               })
               .Add("o|output=", "{RESULT} file with calendar: docx", delegate (string v) {
                   if (v != null)
                   {
                       outputPath = v;
                   }
                   else {
                       messageHandlingStatus.resultUnspecified = true;
                       outputPath = String.Format("{0: yyyyMMdd_mmHHss}.docx", System.DateTime.Now);
                   }
               });
            try {
                List<string> extra = p.Parse(args);
            }
            catch (OptionException e) {
                Console.WriteLine(e.Message);
                p.WriteOptionDescriptions(Console.Out);
            }
            
            if (messageHandlingStatus.show_help) { //user requested for help
                DisplayHelpMessage(p);
                return;
            }
            else {
                if ((firstDate == null) || (secondDate == null)) {
                    System.Console.WriteLine("Explicitly specify both dates");
                    p.WriteOptionDescriptions(Console.Out);
                    return;
                } else
                if (messageHandlingStatus.templateExpected == true ) {
                    System.Console.WriteLine("Explicitly specify path to template");
                    p.WriteOptionDescriptions(Console.Out);
                    return;
                }
            }
            if (messageHandlingStatus.resultUnspecified) {
                System.Console.WriteLine("You have not specified path to result...");
            }
            System.Console.WriteLine("Writing to file {0}", outputPath);
        }

        private static void DisplayHelpMessage(OptionSet p)
        {
            System.Console.WriteLine("== A calendar app. Construct tabletop calendar ==");
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }
    }
}
