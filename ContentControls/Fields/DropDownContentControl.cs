using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TXTextControl.DocumentServer.Fields
{
    public class DropDownListContentControl : ContentControlFieldAdapter
    {
        /*-------------------------------------------------------------------------------------------------------
        ** Fields
        **-----------------------------------------------------------------------------------------------------*/

        private List<DropDownListItem> _lstListItems = new List<DropDownListItem>();
        
        /*-------------------------------------------------------------------------------------------------------
        ** Constructors
        **-----------------------------------------------------------------------------------------------------*/

        public DropDownListContentControl(ApplicationField ApplicationField)
            : base(ApplicationField)
        {
            _lstListItems = new List<DropDownListItem>();

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(ApplicationField.Parameters[0]);

            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xml.NameTable);
            namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNodeList xnList = xml.SelectNodes("/w:sdt/w:sdtPr/w:dropDownList", namespaceManager);

            foreach (XmlNode xn in xnList)
            {
                foreach (XmlNode subnode in xn.ChildNodes)
                {
                    switch (subnode.LocalName)
                    {
                        case "listItem":
                            DropDownListItem listItem = new DropDownListItem();
                            if (subnode.Attributes["w:value"] != null)
                                listItem.Value = subnode.Attributes["w:value"].Value;

                            if (subnode.Attributes["w:displayText"] != null)
                                listItem.DisplayText = subnode.Attributes["w:displayText"].Value;

                            _lstListItems.Add(listItem);
                            break;
                    }
                }
            }
        }

        public DropDownListContentControl()
            : base()
        {
            
        }

        /*-------------------------------------------------------------------------------------------------------
        ** Properties
        **-----------------------------------------------------------------------------------------------------*/

        public List<DropDownListItem> ListItems
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

            System.Xml.XmlElement dropDownElement = xml.CreateElement("w:dropDownList", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            foreach (DropDownListItem item in this.ListItems)
            {
                System.Xml.XmlElement dropDownListItemElement = xml.CreateElement("w:listItem", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                dropDownListItemElement.SetAttribute("displayText", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", item.DisplayText);
                dropDownListItemElement.SetAttribute("value", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", item.Value);

                dropDownElement.AppendChild(dropDownListItemElement);
            }

            xn.AppendChild(dropDownElement);

            this.ApplicationField.Parameters =
                new string[] { xml.OuterXml };
        }
    }

    /*-------------------------------------------------------------------------------------------------------
    ** Helpers
    **-----------------------------------------------------------------------------------------------------*/

    public class DropDownListItem
    {
        public string DisplayText { get; set; }
        public string Value { get; set; }

        public DropDownListItem(string DisplayText, string Value)
        {
            this.DisplayText = DisplayText;
            this.Value = Value;
        }

        public DropDownListItem()
        {

        }
    }
}
