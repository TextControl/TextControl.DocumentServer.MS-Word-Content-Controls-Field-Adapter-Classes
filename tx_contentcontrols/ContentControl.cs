using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using TXTextControl;

namespace tx_contentcontrols
{
    public abstract class ContentControlFieldAdapter
    {
        public ApplicationField ApplicationField { get; set; }
        public string Title { get; set; }
        public string Tag { get; set; }
        public long Id { get; set; }
        public bool ContentDeletable { get; set; }
        public bool ContentEditable { get; set; }
        public string Type { get; set; }
        public List<DocPart> Placeholder { get; set; }

        public static Type GetContentControlType(ApplicationField ApplicationField)
        {
            if (ApplicationField.TypeName != "SDTRUN" && ApplicationField.TypeName != "SDTBLOCK")
                return null;

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(ApplicationField.Parameters[0]);

            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xml.NameTable);
            namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNodeList xnList = xml.SelectNodes("/w:sdt/w:sdtPr", namespaceManager);

            foreach (XmlNode xn in xnList)
            {
                foreach (XmlNode subnode in xn.ChildNodes)
                {
                    switch (subnode.LocalName)
                    {
                        case "text":
                            return typeof(PlainTextContentControl);
                        case "checkbox":
                            return typeof(CheckBoxContentControl);
                        case "comboBox":
                            return typeof(ComboBoxContentControl);
                        case "dropDownList":
                            return typeof(DropDownListContentControl);
                        case "date":
                            return typeof(DateContentControl);
                    }
                }

                return typeof(RichTextContentControl);
            }

            return null;
        }

        public ContentControlFieldAdapter(ApplicationField ApplicationField)
        {
            this.ApplicationField = ApplicationField;
            GetParameters();
        }

        public XmlDocument GetXmlBaseStructure()
        {
            XmlDocument xml = new XmlDocument();

            System.Xml.XmlElement SDT = xml.CreateElement("w:sdt", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            SDT.SetAttribute("xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            SDT.SetAttribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            SDT.SetAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
            SDT.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            SDT.SetAttribute("xmlns:v", "urn:schemas-microsoft-com:vml");
            SDT.SetAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            SDT.SetAttribute("xmlns:w10", "urn:schemas-microsoft-com:office:word");
            SDT.SetAttribute("xmlns:w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            SDT.SetAttribute("xmlns:w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            SDT.SetAttribute("xmlns:wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            SDT.SetAttribute("xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            SDT.SetAttribute("xmlns:wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            SDT.SetAttribute("xmlns:wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            SDT.SetAttribute("xmlns:wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            SDT.SetAttribute("xmlns:wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            SDT.SetAttribute("xmlns:wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            XmlNode SDTNode = xml.AppendChild(SDT);

            System.Xml.XmlElement SDTPR = xml.CreateElement("w:sdtPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            XmlNode SDTPRNode = SDTNode.AppendChild(SDTPR);

            System.Xml.XmlElement alias = xml.CreateElement("w:alias", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            alias.SetAttribute("w:val", this.Title);

            System.Xml.XmlElement tag = xml.CreateElement("w:tag", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            tag.SetAttribute("w:val", this.Tag);

            System.Xml.XmlElement id = xml.CreateElement("w:id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            id.SetAttribute("w:val", this.Id.ToString());

            System.Xml.XmlElement lockElement = xml.CreateElement("w:lock", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            lockElement.SetAttribute("w:val", "sdtContentLocked");

            System.Xml.XmlElement lockElement2 = xml.CreateElement("w:lock", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            lockElement2.SetAttribute("w:val", "ContentLocked");

            System.Xml.XmlElement placeholder = xml.CreateElement("w:placeholder", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            foreach (DocPart part in this.Placeholder)
            {
                System.Xml.XmlElement partElement = xml.CreateElement("w:docPart", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                partElement.SetAttribute("w:val", part.Value);

                placeholder.AppendChild(partElement);
            }

            System.Xml.XmlElement plaintext = xml.CreateElement("w:text", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SDTPRNode.AppendChild(alias);
            SDTPRNode.AppendChild(tag);
            SDTPRNode.AppendChild(id);
            SDTPRNode.AppendChild(placeholder);

            if (this.ContentDeletable == false)
                SDTPRNode.AppendChild(lockElement);
            if (this.ContentEditable == false)
                SDTPRNode.AppendChild(lockElement2);
            if (this.Type == "PLAINTEXT")
                SDTPRNode.AppendChild(plaintext);
            
            System.Xml.XmlElement SDTENDPR = xml.CreateElement("w:sdtEndPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            SDTNode.AppendChild(SDTENDPR);

            return xml;
        }

        protected void GetParameters()
        {
            this.ContentDeletable = true;
            this.ContentEditable = true;

            if (ApplicationField.TypeName != "SDTRUN" && ApplicationField.TypeName != "SDTBLOCK")
                return;

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(ApplicationField.Parameters[0]);

            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xml.NameTable);
            namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNodeList xnList = xml.SelectNodes("/w:sdt/w:sdtPr", namespaceManager);

            foreach (XmlNode xn in xnList)
            {
                foreach (XmlNode subnode in xn.ChildNodes)
                {
                    switch (subnode.LocalName)
                    {
                        case "alias":
                            this.Title = subnode.Attributes["w:val"].Value;
                            break;
                        case "tag":
                            this.Tag = subnode.Attributes["w:val"].Value;
                            break;
                        case "id":
                            this.Id = Convert.ToInt32(subnode.Attributes["w:val"].Value);
                            break;
                        case "lock":
                            if (subnode.Attributes["w:val"].Value == "sdtContentLocked")
                                this.ContentDeletable = false;
                            else
                                this.ContentEditable = false;
                            break;
                        case "text":
                            this.Type = "PLAINTEXT";
                            break;
                        case "checkbox":
                            this.Type = "CHECKBOX";
                            break;
                        case "comboBox":
                            this.Type = "COMBOBOX";
                            break;
                        case "dropDownList":
                            this.Type = "DROPDOWNLIST";
                            break;
                        case "date":
                            this.Type = "DATE";
                            break;
                        case "placeholder":
                            this.Placeholder = new List<DocPart>();   
                         
                            XmlNode placeholderXn = xml.SelectSingleNode("/w:sdt/w:sdtPr/w:placeholder", namespaceManager);

                            foreach (XmlNode docPart in placeholderXn.ChildNodes)
                            {
                                this.Placeholder.Add(new DocPart(docPart.Attributes["w:val"].Value));
                            }

                            break;
                    }
                }
            }
        }
    }    

    public class RichTextContentControl : ContentControlFieldAdapter
    {
        public RichTextContentControl(ApplicationField ApplicationField)
            : base(ApplicationField)
        {
            
        }

        public string Text
        {
            get { return this.ApplicationField.Text; }
            set
            {
                this.ApplicationField.Text = value;
            }
        }

        public void SetParameters()
        {
            XmlDocument xml = this.GetXmlBaseStructure();
            this.ApplicationField.Parameters[0] = xml.OuterXml;
        }
    }

    public class PlainTextContentControl : ContentControlFieldAdapter
    {
        public PlainTextContentControl(ApplicationField ApplicationField)
            : base(ApplicationField)
        {
            
        }

        public string Text
        {
            get { return this.ApplicationField.Text; }
            set
            {
                this.ApplicationField.Text = value;
            }
        }

        public void SetParameters()
        {
            XmlDocument xml = this.GetXmlBaseStructure();
            this.ApplicationField.Parameters[0] = xml.OuterXml;
        }
    }

    public class CheckBoxContentControl : ContentControlFieldAdapter
    {
        private bool m_bChecked = false;

        public bool Checked
        {
            get { return m_bChecked; }
            set
            {
                m_bChecked = value;
                this.ApplicationField.Text = m_bChecked ? "☒\r\n" : "☐\r\n";
            }
        }

        public void SetParameters()
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

            this.ApplicationField.Parameters[0] = xml.OuterXml;
        }

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
                            this.Checked = ((Convert.ToInt32(subnode.Attributes["w14:val"].Value) != 0));
                            break;
                    }
                }
            }
        }
    }

    public class ComboBoxContentControl : ContentControlFieldAdapter
    {
        public List<ComboBoxListItem> ListItems { get; set; }

        public ComboBoxContentControl(ApplicationField ApplicationField)
            : base(ApplicationField)
        {
            ListItems = new List<ComboBoxListItem>();

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
                            listItem.DisplayText = subnode.Attributes["w:displayText"].Value;
                            listItem.Value = subnode.Attributes["w:value"].Value;

                            ListItems.Add(listItem);
                            break;
                    }
                }
            }
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

    public class DropDownListContentControl : ContentControlFieldAdapter
    {
        public List<DropDownListItem> ListItems { get; set; }

        public DropDownListContentControl(ApplicationField ApplicationField)
            : base(ApplicationField)
        {
            ListItems = new List<DropDownListItem>();

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
                            listItem.Value = subnode.Attributes["w:value"].Value;

                            ListItems.Add(listItem);
                            break;
                    }
                }
            }
        }
    }

    public class DocPart
    {
        public string Value { get; set; }

        public DocPart(string Value)
        {
            this.Value = Value;
        }

        public DocPart()
        {

        }
    }

    public class DropDownListItem
    {
        public string Value { get; set; }
        
        public DropDownListItem(string Value)
        {
            this.Value = Value;
        }

        public DropDownListItem()
        {
            
        }
    }

    public class DateContentControl : ContentControlFieldAdapter
    {
        private DateTime m_dtDate;

        public DateTime Date
        {
            get { return m_dtDate; }
            set
            {
                m_dtDate = value;
                this.ApplicationField.Text = m_dtDate.ToString(this.DateFormat);
            }
        }
        public string DateFormat { get; set; }
        public string LanguageID { get; set; }
        public string StoreMappedDataAs { get; set; }
        public string Calendar { get; set; }

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
                    m_dtDate = Convert.ToDateTime(xn.Attributes["w:fullDate"].Value);

                foreach (XmlNode subnode in xn.ChildNodes)
                {
                    switch (subnode.LocalName)
                    {
                        case "dateFormat":
                            this.DateFormat = subnode.Attributes["w:val"].Value;
                            break;
                        case "lid":
                            this.LanguageID = subnode.Attributes["w:val"].Value;
                            break;
                        case "storeMappedDataAs":
                            this.StoreMappedDataAs = subnode.Attributes["w:val"].Value;
                            break;
                        case "calendar":
                            this.Calendar = subnode.Attributes["w:val"].Value;
                            break;
                    }
                }
            }

            this.Date = m_dtDate;
        }
    }

    
}
