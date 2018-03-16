using System;
using System.Collections.Generic;
using System.Drawing;
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
            pageWidth = CodeHandling.cmToPoints(29.7f);
            pageHeight = CodeHandling.cmToPoints(21.0f);
            numcolsTable = 3; numrowsTable = 2;
        }
    }
    public class CodeHandling
    {
        private System.DateTime date1; private System.DateTime date2;
        public void constructDocument(System.DateTime in_date1, System.DateTime in_date2) {
            if (in_date1 > in_date2) {
                date1 = in_date2; date2 = in_date1;
            }
            else {
                date1 = in_date1; date2 = in_date2;
            }
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

            System.DateTime theCurrentDate = date1;
            do
            {
                document.PageHeight = in_documentArgs.pageHeight;
                document.PageWidth = in_documentArgs.pageWidth;
                document.MarginBottom = cmToPoints(1.0f); document.MarginTop = cmToPoints(1.0f); document.MarginLeft = cmToPoints(1.0f); document.MarginRight = cmToPoints(1.0f);
                Table insertedTable = document.InsertTable(in_documentArgs.numrowsTable, in_documentArgs.numcolsTable);
                Border b = new Border(BorderStyle.Tcbs_single, BorderSize.one, 0, Color.Blue);

                // Set the tables Top, Bottom, Left and Right Borders to b.
                insertedTable.SetBorder(TableBorderType.Top, b);
                insertedTable.SetBorder(TableBorderType.Bottom, b);
                insertedTable.SetBorder(TableBorderType.Left, b);
                insertedTable.SetBorder(TableBorderType.Right, b);
                insertedTable.SetBorder(TableBorderType.InsideH, b);
                insertedTable.SetBorder(TableBorderType.InsideV, b);
                byte currentRow = 0; byte currentCol = 0;
                while ((currentRow < in_documentArgs.numrowsTable) && (currentCol < in_documentArgs.numcolsTable)&& (theCurrentDate <= date2)) {
                    insertedTable.Rows[currentRow].Cells[currentCol].InsertParagraph(String.Format("{0:yyyy MMMM dd}", theCurrentDate));
                    theCurrentDate = theCurrentDate.AddDays(1.0);
                    currentCol++;
                    if ((currentCol >= in_documentArgs.numcolsTable)&&(currentRow<in_documentArgs.numrowsTable-1)) {
                        currentCol = 0; currentRow++;
                    }
                }
                if (theCurrentDate <= date2) {
                    document.InsertSectionPageBreak();
                }
            } while (theCurrentDate <= date2);
            
            
            document.InsertTable(in_documentArgs.numrowsTable, in_documentArgs.numcolsTable);
            document.Save();
        }
    }
}
