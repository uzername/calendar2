using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
// https://www.codeproject.com/Articles/483055/XML-Serialization-and-Deserialization-Part-1
// https://www.codeproject.com/Articles/487571/XML-Serialization-and-Deserialization-Part-2
namespace XMLprocessing {
	public class importantdays {
		[XmlElement("day")]
		public List<importantday> importantdayList = new List<importantday>();
		
	}
	public class importantday {
		[XmlAttribute("type")]
		public String typeOfDay;
		[XmlElement("date")]
		public String date;
		[XmlElement("description")]
		public String description;
	}
    public class MainXMLprocessor {
		public importantdays loadImportantDaysListFromFile(String in_pathToXMLfile) {
            XmlSerializer deserializer = new XmlSerializer(typeof(importantdays));
            TextReader reader = null ; importantdays XmlData;
            try
            {
                reader = new StreamReader(in_pathToXMLfile);
                object obj = deserializer.Deserialize(reader);
                XmlData = (importantdays)obj;
            } finally
            {
                reader.Close();
            }    
			return XmlData;
        }
        /// <summary>
        /// we get just raw array. clasterize it by dates
        /// </summary>
        /// <param name="in_importantdays"></param>
        /// <returns></returns>
	    public Dictionary<System.DateTime, List<importantday>> getDictionaryForProcessing (importantdays in_importantdays) {
	        Dictionary<System.DateTime, List<importantday>> valueToReturn = new Dictionary<System.DateTime, List<importantday>>();
            foreach (importantday item in in_importantdays.importantdayList) {
                DateTime foundDate = DateTime.Parse(item.date);
                if (valueToReturn.ContainsKey(foundDate)) {
                    valueToReturn[foundDate].Add(item);
                }
                else {
                    List<importantday> listToUse = new List<importantday>(); listToUse.Add(item);
                    valueToReturn.Add(foundDate, listToUse);
                }

            }
		    return valueToReturn;
	    }
	}
}
