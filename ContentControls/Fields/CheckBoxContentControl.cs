using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TXTextControl.DocumentServer.Fields
{
    public class CheckBoxContentControl : ContentControlFieldAdapter
    { 
        /*-------------------------------------------------------------------------------------------------------
        ** Fields
        **-----------------------------------------------------------------------------------------------------*/
        
        private bool _bChecked;

        /*-------------------------------------------------------------------------------------------------------
        ** Constructors
        **-----------------------------------------------------------------------------------------------------*/
        
        public CheckBoxContentControl(ApplicationField ApplicationField)
            : base(ApplicationField)
        {
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(ApplicationField.Parameters[0]);

            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xml.NameTable);
            namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            namespaceManager.AddNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            XmlNodeList xnList = xml.SelectNodes("/w:sdt/w:sdtPr/w14:checkbox", namespaceManager);

            foreach (XmlNode xn in xnList)
            {
                foreach (XmlNode subnode in xn.ChildNodes)
                {
                    switch (subnode.LocalName)
                    {
                        case "checked":
                            if (subnode.Attributes["w14:val"] != null)
                                _bChecked = ((Convert.ToInt32(subnode.Attributes["w14:val"].Value) != 0));
                            break;
                    }
                }
            }
        }

        public CheckBoxContentControl()
            : base()
        {
            
        }

        /*-------------------------------------------------------------------------------------------------------
        ** Properties
        **-----------------------------------------------------------------------------------------------------*/
        public bool Checked
        {
            get { return _bChecked; }
            set
            {
                _bChecked = value;
                this.ApplicationField.Text = _bChecked ? "☒\r\n" : "☐\r\n";
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

            System.Xml.XmlElement checkBoxElement = xml.CreateElement("w14:checkbox", "http://schemas.microsoft.com/office/word/2010/wordml");
            System.Xml.XmlElement checkedElement = xml.CreateElement("w14:checked", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            int i = this.Checked ? 1 : 0;
            checkedElement.SetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", i.ToString());

            checkBoxElement.AppendChild(checkedElement);
            xn.AppendChild(checkBoxElement);

            this.ApplicationField.Parameters = 
                new string[] { xml.OuterXml };
        }

        
    }
}
