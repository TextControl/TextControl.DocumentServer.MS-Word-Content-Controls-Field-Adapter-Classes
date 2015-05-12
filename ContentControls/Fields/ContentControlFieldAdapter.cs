using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TXTextControl.DocumentServer.Fields
{
    public abstract class ContentControlFieldAdapter
    {
        /*-------------------------------------------------------------------------------------------------------
        ** Fields
        **-----------------------------------------------------------------------------------------------------*/

        private ApplicationField _afApplicationField;
        private string _strTitle = "";
        private string _strTag = "";
        private string _strType;
        private long _lId;
        private bool _bContentDeletable = true;
        private bool _bContentEditable = true;
        private List<DocPart> _lstPlaceholder = new List<DocPart>();

        /*-------------------------------------------------------------------------------------------------------
        ** Constructors
        **-----------------------------------------------------------------------------------------------------*/

        public ContentControlFieldAdapter(ApplicationField ApplicationField)
        {
            _afApplicationField = ApplicationField;
            GetParameters();
        }

        public ContentControlFieldAdapter()
        {
            this.ApplicationField = new ApplicationField(ApplicationFieldFormat.MSWord, "SDTBLOCK", "Field", new string[] { this.GetXmlBaseStructure().OuterXml });
        }

        /*-------------------------------------------------------------------------------------------------------
        ** Properties
        **-----------------------------------------------------------------------------------------------------*/

        public ApplicationField ApplicationField
        {
            get
            {
                return _afApplicationField;
            }
            set
            {
                _afApplicationField = value;
                SetParameters();
            }
        }

        public string Title
        {
            get { return _strTitle; }
            set
            {
                _strTitle = value;
                SetParameters();
            }
        }

        public string Tag
        {
            get { return _strTag; }
            set
            {
                _strTag = value;
                SetParameters();
            }
        }

        public long Id
        {
            get { return _lId; }
            set
            {
                _lId = value;
                SetParameters();
            }
        }

        public bool ContentDeletable
        {
            get { return _bContentDeletable; }
            set
            {
                _bContentDeletable = value;
                SetParameters();
            }
        }

        public bool ContentEditable
        {
            get { return _bContentEditable; }
            set
            {
                _bContentEditable = value;
                SetParameters();
            }
        }

        public string Type
        {
            get { return _strType; }
            set
            {
                _strType = value;
                SetParameters();
            }
        }

        public List<DocPart> Placeholder
        {
            get { return _lstPlaceholder; }
            set
            {
                _lstPlaceholder = value;
                SetParameters();
            }
        }

        /*-------------------------------------------------------------------------------------------------------
        ** Methods
        **-----------------------------------------------------------------------------------------------------*/

        protected abstract void SetParameters();

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

            if (this.Placeholder != null)
            {
                System.Xml.XmlElement placeholder = xml.CreateElement("w:placeholder", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                foreach (DocPart part in this.Placeholder)
                {
                    System.Xml.XmlElement partElement = xml.CreateElement("w:docPart", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    partElement.SetAttribute("w:val", part.Value);

                    placeholder.AppendChild(partElement);
                }

                SDTPRNode.AppendChild(placeholder);
            }

            System.Xml.XmlElement plaintext = xml.CreateElement("w:text", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SDTPRNode.AppendChild(alias);
            SDTPRNode.AppendChild(tag);
            SDTPRNode.AppendChild(id);

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
            if (ApplicationField.TypeName != "SDTRUN" && ApplicationField.TypeName != "SDTBLOCK" && ApplicationField.Parameters[0] == null)
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
                            if(subnode.Attributes["w:val"] != null)
                                _strTitle = subnode.Attributes["w:val"].Value;
                            break;
                        case "tag":
                            if (subnode.Attributes["w:val"] != null)
                                _strTag = subnode.Attributes["w:val"].Value;
                            break;
                        case "id":
                            if (subnode.Attributes["w:val"] != null)
                                _lId = Convert.ToInt32(subnode.Attributes["w:val"].Value);
                            break;
                        case "lock":
                            if (subnode.Attributes["w:val"] != null)
                            {
                                if (subnode.Attributes["w:val"].Value == "sdtContentLocked")
                                    _bContentDeletable = false;
                                else
                                    _bContentEditable = false;
                            }
                            break;
                        case "text":
                            _strType = "PLAINTEXT";
                            break;
                        case "checkbox":
                            _strType = "CHECKBOX";
                            break;
                        case "comboBox":
                            _strType = "COMBOBOX";
                            break;
                        case "dropDownList":
                            _strType = "DROPDOWNLIST";
                            break;
                        case "date":
                            _strType = "DATE";
                            break;
                        case "placeholder":
                            _lstPlaceholder = new List<DocPart>();

                            XmlNode placeholderXn = xml.SelectSingleNode("/w:sdt/w:sdtPr/w:placeholder", namespaceManager);

                            foreach (XmlNode docPart in placeholderXn.ChildNodes)
                            {
                                if (docPart.Attributes["w:val"] != null)
                                    _lstPlaceholder.Add(new DocPart(docPart.Attributes["w:val"].Value));
                            }

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
}
