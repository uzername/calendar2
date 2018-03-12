using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xceed.Words.NET;

namespace ConsoleAppCalendar
{
    public class CodeHandling
    {
        public void constructDocument() {
            DocX document = DocX.Create("Indentation.docx");
            // Add a title.
            document.InsertParagraph("Paragraph indentation").FontSize(15d).SpacingAfter(50d).Alignment = Alignment.center;

            // Set a smaller page width. (in points)
            document.PageWidth = 250f;
            document.Save();
        }
        /// <summary>
        /// transform value in cm to value in points. 1pt = 1/72 of inch, inch = 72pt; 1cm=72/2.54 pt; 1pt = 2.54/72 cm
        /// https://www.asknumbers.com/CentimetersToPointsConversion.aspx
        /// </summary>
        /// <param name="in_cmValue"></param>
        /// <returns></returns>
        public static float cmToPoints(float in_cmValue) {
            return in_cmValue / ((float)2.54 / (float)72.0);
        }


    }
}
