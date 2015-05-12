using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TXTextControl.DocumentServer.Fields
{
    public class ComboBoxContentControl : ContentControlFieldAdapter
    {
        /*-------------------------------------------------------------------------------------------------------
        ** Fields
        **-----------------------------------------------------------------------------------------------------*/
        
        private List<ComboBoxListItem> _lstListItems = new List<ComboBoxListItem>();

        /*-------------------------------------------------------------------------------------------------------
        ** Constructors
        **-----------------------------------------------------------------------------------------------------*/
        public ComboBoxContentControl(ApplicationField ApplicationField)
            : base(ApplicationField)
        {
            _lstListItems = new List<ComboBoxListItem>();

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(ApplicationField.Parameters[0]);

            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xml.NameTable);
            namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNodeList xnList = xml.SelectNodes("/w:sdt/w:sdtPr/w:comboBox", namespaceManager);

            foreach (XmlNode xn in xnList)
            {
                foreach (XmlNode subnode in xn.ChildNodes)
                {
                    switch (subnode.LocalName)
                    {
                        case "listItem":
                            ComboBoxListItem listItem = new ComboBoxListItem();
                            if (subnode.Attributes["w:displayText"] != null)
                                listItem.DisplayText = subnode.Attributes["w:displayText"].Value;
                            if (subnode.Attributes["w:displayText"] != null)
                                listItem.Value = subnode.Attributes["w:value"].Value;

                            _lstListItems.Add(listItem);
                            break;
                    }
                }
            }
        }

        public ComboBoxContentControl()
            : base()
        {
            
        }

        /*-------------------------------------------------------------------------------------------------------
        ** Properties
        **-----------------------------------------------------------------------------------------------------*/
        
        public List<ComboBoxListItem> ListItems 
        {
            get { return _lstListItems; }
            set
            {
                _lstListItems = value;
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

            System.Xml.XmlElement comboBoxElement = xml.CreateElement("w:comboBox", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            foreach (ComboBoxListItem item in this.ListItems)
            {
                System.Xml.XmlElement listItemElement = xml.CreateElement("w:listItem", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                listItemElement.SetAttribute("displayText", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", item.DisplayText);
                listItemElement.SetAttribute("value", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", item.Value);

                comboBoxElement.AppendChild(listItemElement);
            }

            xn.AppendChild(comboBoxElement);

            this.ApplicationField.Parameters =
                new string[] { xml.OuterXml };
        }
    }

    public class ComboBoxListItem
    {
        public string DisplayText { get; set; }
        public string Value { get; set; }

        public ComboBoxListItem()
        {

        }

        public ComboBoxListItem(string DisplayText, string Value)
        {
            this.DisplayText = DisplayText;
            this.Value = Value;
        }
    }
}
