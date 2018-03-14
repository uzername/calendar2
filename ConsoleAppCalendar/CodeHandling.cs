using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xceed.Words.NET;

namespace ConsoleAppCalendar
{
    public class GenericDocumentParameters {
        public String documentName { get; set; }
        public float pageWidth { get; set; }
        public float pageHeight { get; set; }
        public System.Byte numcolsTable { get; set; }
        public System.Byte numrowsTable { get; set; }
        public GenericDocumentParameters()
        {
            documentName = "MyDocument.docx";
            pageWidth = CodeHandling.cmToPoints(21.0f);
            pageHeight = CodeHandling.cmToPoints(29.7f);
            numcolsTable = 3; numrowsTable = 2;
        }
    }
    public class CodeHandling
    {
        public void constructDocument() {
            
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
        public void testCreateDocument() {
            DocX document = DocX.Create("Indentation.docx");
            // Add a title.
            document.InsertParagraph("Paragraph indentation").FontSize(15d).SpacingAfter(50d).Alignment = Alignment.center;
            // Set a smaller page width. (in points)
            document.PageWidth = 250f;
            document.Save();
        }
        public void fromvariableCreateDocument(GenericDocumentParameters in_documentArgs) {
            DocX document = DocX.Create(in_documentArgs.documentName);
            document.PageHeight = in_documentArgs.pageHeight;
            document.PageWidth = in_documentArgs.pageWidth;
            document.InsertTable(in_documentArgs.numrowsTable, in_documentArgs.numcolsTable);
            document.InsertSectionPageBreak();
            document.InsertTable(in_documentArgs.numrowsTable, in_documentArgs.numcolsTable);
            document.Save();
        }
    }
}
