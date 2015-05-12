using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TXTextControl.DocumentServer.Fields
{
    public class DateContentControl : ContentControlFieldAdapter
    {
        /*-------------------------------------------------------------------------------------------------------
        ** Fields
        **-----------------------------------------------------------------------------------------------------*/

        private DateTime _dtDate;
        private string _strDateFormat = "M/d/yyyy";
        private string _strLanguageID = "en-US";
        private string _strStoreMappedDataAs = "dateTime";
        private string _strCalendar = "gregorian";

        /*-------------------------------------------------------------------------------------------------------
        ** Constructors
        **-----------------------------------------------------------------------------------------------------*/

        public DateContentControl(ApplicationField ApplicationField)
            : base(ApplicationField)
        {
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(ApplicationField.Parameters[0]);

            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xml.NameTable);
            namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNodeList xnList = xml.SelectNodes("/w:sdt/w:sdtPr/w:date", namespaceManager);

            foreach (XmlNode xn in xnList)
            {
                if (xn.Attributes["w:fullDate"].Value != null)
                    _dtDate = Convert.ToDateTime(xn.Attributes["w:fullDate"].Value);

                foreach (XmlNode subnode in xn.ChildNodes)
                {
                    switch (subnode.LocalName)
                    {
                        case "dateFormat":
                            if (subnode.Attributes["w:val"] != null)
                                _strDateFormat = subnode.Attributes["w:val"].Value;
                            break;
                        case "lid":
                            if (subnode.Attributes["w:val"] != null)
                                _strLanguageID = subnode.Attributes["w:val"].Value;
                            break;
                        case "storeMappedDataAs":
                            if (subnode.Attributes["w:val"] != null)
                                _strStoreMappedDataAs = subnode.Attributes["w:val"].Value;
                            break;
                        case "calendar":
                            if (subnode.Attributes["w:val"] != null)
                                _strCalendar = subnode.Attributes["w:val"].Value;
                            break;
                    }
                }
            }

            this.Date = _dtDate;
        }

        public DateContentControl()
            : base()
        {
            
        }

        /*-------------------------------------------------------------------------------------------------------
        ** Properties
        **-----------------------------------------------------------------------------------------------------*/

        public DateTime Date
        {
            get { return _dtDate; }
            set
            {
                _dtDate = value;
                this.ApplicationField.Text = _dtDate.ToString(this.DateFormat);
                SetParameters();
            }
        }

        public string DateFormat
        {
            get { return _strDateFormat; }
            set
            {
                _strDateFormat = value;
                SetParameters();
            }
        }

        public string LanguageID
        {
            get { return _strLanguageID; }
            set
            {
                _strLanguageID = value;
                SetParameters();
            }
        }

        public string StoreMappedDataAs
        {
            get { return _strStoreMappedDataAs; }
            set
            {
                _strStoreMappedDataAs = value;
                SetParameters();
            }
        }

        public string Calendar
        {
            get { return _strCalendar; }
            set
            {
                _strCalendar = value;
                SetParameters();
            }
        }

        /*-------------------------------------------------------------------------------------------------------
        ** Methods
        **-----------------------------------------------------------------------------------------------------*/

        protected override void SetParameters()
        {
            XmlDocument xml = this.GetXmlBaseStructure();

            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xml.NameTable);
            namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNode xn = xml.SelectSingleNode("/w:sdt/w:sdtPr", namespaceManager);

            System.Xml.XmlElement dateElement = xml.CreateElement("w:date", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            if (this.Date != null)
                dateElement.SetAttribute("fullDate", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", this.Date.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ssZ"));

            System.Xml.XmlElement dateFormatElement = xml.CreateElement("w:dateFormat", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            if (this.DateFormat != null)
                dateFormatElement.SetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", this.DateFormat);

            System.Xml.XmlElement lidElement = xml.CreateElement("w:lid", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            if (this.LanguageID != null)
                lidElement.SetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", this.LanguageID);

            System.Xml.XmlElement storeMappedDataAsElement = xml.CreateElement("w:storeMappedDataAs", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            if (this.StoreMappedDataAs != null)
                storeMappedDataAsElement.SetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", this.StoreMappedDataAs);

            System.Xml.XmlElement calendarDataAsElement = xml.CreateElement("w:calendar", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            if (this.Calendar != null)
                calendarDataAsElement.SetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", this.Calendar);

            dateElement.AppendChild(dateFormatElement);
            dateElement.AppendChild(lidElement);
            dateElement.AppendChild(storeMappedDataAsElement);
            dateElement.AppendChild(calendarDataAsElement);

            xn.AppendChild(dateElement);

            this.ApplicationField.Parameters =
                new string[] { xml.OuterXml };
        }

        
    }
}
