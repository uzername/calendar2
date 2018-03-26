using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SunSetRiseLib;
using Xceed.Words.NET;
using XMLprocessing;

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
            numcolsTable = 3; numrowsTable = 4;
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
        /// <summary>
        /// http://pointofint.blogspot.com/2014/06/sunrise-and-sunset-in-c.html
        /// </summary>
        /// <param name="latitude"> More than 0 if northern lat</param>
        /// <param name="longitude"> More than 0 if eastern long</param>
        /// <returns></returns>
        public static Tuple<string,string> getSunsetAndSunRise(Boolean useDayLightSaveTime, int timezoneCorr, System.DateTime date, double latitude, double longitude)
        {
            double JD = Util.calcJD(date);  //OR   JD = Util.calcJD(2014, 6, 1);
            double sunRise = Util.calcSunRiseUTC(JD, latitude, longitude);
            double sunSet = Util.calcSunSetUTC(JD, latitude, longitude);
            string sunrisetimesrtr = Util.getTimeString(sunRise, timezoneCorr, JD, useDayLightSaveTime);
            string sunsettimesrtr = Util.getTimeString(sunSet, timezoneCorr, JD, useDayLightSaveTime);
            return new Tuple<string, string>(sunrisetimesrtr, sunsettimesrtr);
        }
        

        /// <summary>
        /// convert cm to pixels by multiplying points value on 2/3
        /// </summary>
        /// <param name="in_cmValue"></param>
        /// <returns></returns>
        public static float cmToPixels(float in_cmValue) {
            return cmToPoints(in_cmValue) * ((float)2.0 / (float)3.0);
        }
        public void testCreateDocument() {
            DocX document = DocX.Create("Indentation.docx");
            // Add a title.
            document.InsertParagraph("Paragraph indentation").FontSize(15d).SpacingAfter(50d).Alignment = Alignment.center;
            // Set a smaller page width. (in points)
            document.PageWidth = 250f;
            document.Save();
        }
        public void fromvariableCreateDocument(GenericDocumentParameters in_documentArgs, String obtainedPath) {
            //prepare important days list
            MainXMLprocessor importantDaysProcessor = new MainXMLprocessor();
            importantdays alldaysListRaw = importantDaysProcessor.loadImportantDaysListFromFile(obtainedPath);
            Dictionary<System.DateTime, List<importantday>> importantDaysProcessed = importantDaysProcessor.getDictionaryForProcessing(alldaysListRaw);
            //=========
                DocX document = DocX.Create(in_documentArgs.documentName);
            // https://stackoverflow.com/a/90699/
            ConsoleAppCalendar.Properties.Resources.sunrise1f305.Save("sunrisetmp.png");
            ConsoleAppCalendar.Properties.Resources.sunset1f307.Save("sunsettmp.png");
            // https://xceed.com/wp-content/documentation/xceed-words-for-net/Xceed.Words.NET~Xceed.Words.NET.DocX~AddImage(Stream,String).html
            Xceed.Words.NET.Image imgSunrise = document.AddImage("sunrisetmp.png"); Xceed.Words.NET.Picture picSunrise = imgSunrise.CreatePicture();
                picSunrise.Height = (int) Math.Round ( picSunrise.Height * 0.32 );
                picSunrise.Width = (int)Math.Round( picSunrise.Width * 0.32 );
                Xceed.Words.NET.Image imgSunset = document.AddImage("sunsettmp.png"); Xceed.Words.NET.Picture picSunset = imgSunset.CreatePicture();
            picSunset.Height = (int)Math.Round(picSunset.Height * 0.38);
            picSunset.Width = (int)Math.Round(picSunset.Width * 0.38);
            System.DateTime theCurrentDate = date1;
            do
            {
                document.PageHeight = in_documentArgs.pageHeight;
                document.PageWidth = in_documentArgs.pageWidth;
                float btmMargin = 1.0f; float topMargin = 1.0f; float leftMargin = 1.0f; float rightMargin = 1.0f;
                document.MarginBottom = cmToPoints(btmMargin); document.MarginTop = cmToPoints(topMargin); document.MarginLeft = cmToPoints(leftMargin); document.MarginRight = cmToPoints(rightMargin);
                Table insertedTable = document.InsertTable(in_documentArgs.numrowsTable, in_documentArgs.numcolsTable);
                Border b = new Border(BorderStyle.Tcbs_single, BorderSize.one, 0, Color.Blue);
                //calculate each column width
                float bestColumnWidth = cmToPixels(((float)in_documentArgs.pageWidth - (float)(leftMargin + rightMargin)) / (float)(in_documentArgs.numcolsTable));
                float bestRowHeightPt = (((float)in_documentArgs.pageHeight - (float)(topMargin + btmMargin)) / (float)(in_documentArgs.numrowsTable));
                // Set the tables Top, Bottom, Left and Right Borders to b.
                insertedTable.SetBorder(TableBorderType.Top, b);
                insertedTable.SetBorder(TableBorderType.Bottom, b);
                insertedTable.SetBorder(TableBorderType.Left, b);
                insertedTable.SetBorder(TableBorderType.Right, b);
                insertedTable.SetBorder(TableBorderType.InsideH, b);
                insertedTable.SetBorder(TableBorderType.InsideV, b);
                byte currentRow = 0; byte currentCol = 0;
                while ((currentRow < in_documentArgs.numrowsTable) && (currentCol < in_documentArgs.numcolsTable) && (theCurrentDate <= date2)) {
                    Table internalTable1 = insertedTable.Rows[currentRow].Cells[currentCol].InsertTable(1, 2);
                    internalTable1.SetWidthsPercentage(new float[] { 30.0f, 70.0f }, 100.0f);
                    /*
                        Paragraph yearMonthP = insertedTable.Rows[currentRow].Cells[currentCol].InsertParagraph(String.Format("{0:yyyy, MMMM}", theCurrentDate));
                        Paragraph dayNumberP = insertedTable.Rows[currentRow].Cells[currentCol].InsertParagraph(String.Format("{0:dd}", theCurrentDate));
                        Paragraph weekdayP = insertedTable.Rows[currentRow].Cells[currentCol].InsertParagraph(String.Format("{0:dddd}", theCurrentDate));
                    */
                    Paragraph yearMonthP = internalTable1.Rows[0].Cells[1].InsertParagraph(String.Format("{0:yyyy, MMMM}", theCurrentDate));
                    Paragraph dayNumberP = internalTable1.Rows[0].Cells[1].InsertParagraph(String.Format("{0:dd}", theCurrentDate));
                    Paragraph weekdayP = internalTable1.Rows[0].Cells[1].InsertParagraph(String.Format("{0:dddd}", theCurrentDate));
                    yearMonthP.Alignment = Alignment.center; yearMonthP.Font("Courier New");
                    dayNumberP.Alignment = Alignment.center; dayNumberP.Font("Courier New"); dayNumberP.FontSize(15); dayNumberP.Bold();
                    weekdayP.Alignment = Alignment.center; weekdayP.Font("Courier New");

                    Tuple<string, string> sunTimes = getSunsetAndSunRise(true, 2, theCurrentDate, 49.4444, 32.0597);
                    Paragraph sunriseParagraph = internalTable1.Rows[0].Cells[0].InsertParagraph(String.Format("{0}", sunTimes.Item1));
                        // https://xceed.com/wp-content/documentation/xceed-words-for-net/webframe.html#Xceed.Words.NET~Xceed.Words.NET.Paragraph~InsertPicture.html
                        sunriseParagraph.InsertPicture(picSunrise);
                    Paragraph sunsetParagraph = internalTable1.Rows[0].Cells[0].InsertParagraph(String.Format("{0}", sunTimes.Item2));
                    sunsetParagraph.InsertPicture(picSunset);
                    internalTable1.Rows[0].Cells[1].Paragraphs[0].Remove(false);
                    internalTable1.Rows[0].Cells[0].Paragraphs[0].Remove(false);
                    insertedTable.Rows[currentRow].Cells[currentCol].Paragraphs[0].Remove(false);
                    insertedTable.SetColumnWidth(currentCol, bestColumnWidth);
                    // 100*2/3pt -> 3.53 cm
                    // x  px -> bestRowHeightCm
                    if (importantDaysProcessed.ContainsKey(theCurrentDate))
                    {
                        List<importantday> eventsForCurrentDate = importantDaysProcessed[theCurrentDate];
                        bool holidayDisplayed = false;
                        foreach (importantday specialevent in eventsForCurrentDate)
                        {
                            Paragraph eventParagraph = insertedTable.Rows[currentRow].Cells[currentCol].InsertParagraph("\u25EF "+specialevent.description);
                            eventParagraph.Font("Courier New");
                            if ((specialevent.typeOfDay == "holiday")&&(holidayDisplayed==false)) {
                                holidayDisplayed = true;
                                Paragraph officialWeekendParagraph = internalTable1.Rows[0].Cells[1].InsertParagraph("/вихідний/");
                                officialWeekendParagraph.Alignment = Alignment.center; officialWeekendParagraph.Font("Courier New"); officialWeekendParagraph.FontSize(7.0); officialWeekendParagraph.Bold();
                            }
                        }
                    }

                    insertedTable.Rows[currentRow].Height = Math.Round(bestRowHeightPt * 0.88);

                    theCurrentDate = theCurrentDate.AddDays(1.0);
                    currentCol++;
                    if ((currentCol >= in_documentArgs.numcolsTable) && (currentRow < in_documentArgs.numrowsTable - 1)) {
                        currentCol = 0; currentRow++;
                    }
                }

                if (theCurrentDate <= date2) {
                    //document.InsertSectionPageBreak();
                    insertedTable.InsertPageBreakAfterSelf();
                }
            } while (theCurrentDate <= date2);


            //document.InsertTable(in_documentArgs.numrowsTable, in_documentArgs.numcolsTable);
            document.Save();
            System.IO.File.Delete("sunsettmp.png");
            System.IO.File.Delete("sunrise.png");
        }
        }
    
}
