using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
// https://www.codeproject.com/Articles/483055/XML-Serialization-and-Deserialization-Part-1
// https://www.codeproject.com/Articles/487571/XML-Serialization-and-Deserialization-Part-2
namespace XMLprocessing {
	public class importantdays {
		[XmlElement("importantday")]
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
	    public Dictionary<System.DateTime, List<importantday>> getDictionaryForProcessing (importantdays in_importantdays) {
	        Dictionary<System.DateTime, List<importantday>> valueToReturn = new Dictionary<System.DateTime, List<importantday>>();
		return valueToReturn;
	    }
	}
}
