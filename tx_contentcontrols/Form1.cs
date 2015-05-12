using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TXTextControl;
using TXTextControl.DocumentServer.Fields;

namespace tx_contentcontrols
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void newCheckboxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // create a new CheckBoxContentControl and add it to TextControl
            CheckBoxContentControl checkbox = new CheckBoxContentControl();
            checkbox.Checked = true;

            textControl1.ApplicationFields.Add(checkbox.ApplicationField);
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TXTextControl.LoadSettings ls = new TXTextControl.LoadSettings();
            ls.ApplicationFieldFormat = TXTextControl.ApplicationFieldFormat.MSWord;

            // load a sample document with all supported types of
            // content controls
            textControl1.Load("test.docx", TXTextControl.StreamType.WordprocessingML, ls);

            // loop through all fields
            foreach (ApplicationField field in textControl1.ApplicationFields)
            {
                Type type = ContentControlFieldAdapter.GetContentControlType(field);

                // based on the type, create a new ContentControl object and
                // display some field information in a MessageBox
                switch (type.Name)
                {
                    case "RichTextContentControl":
                        RichTextContentControl rtb = new RichTextContentControl(field);

                        MessageBox.Show("RichTextContentControl:\r\nText: " + rtb.Text + "\r\n" +
                            "Title: " + rtb.Title + "\r\n" +
                            "Tag: " + rtb.Tag + "\r\n");
                        break;

                    case "PlainTextContentControl":
                        PlainTextContentControl ptc = new PlainTextContentControl(field);

                        MessageBox.Show("PlainTextContentControl:\r\nText: " + ptc.Text + "\r\n" +
                            "Title: " + ptc.Title + "\r\n" +
                            "Tag: " + ptc.Tag + "\r\n");
                        break;

                    case "CheckBoxContentControl":
                        CheckBoxContentControl check = new CheckBoxContentControl(field);

                        MessageBox.Show("CheckBoxContentControl:\r\nChecked: " + check.Checked.ToString() + "\r\n" +
                            "Title: " + check.Title + "\r\n" +
                            "Tag: " + check.Tag + "\r\n");
                        break;

                    case "ComboBoxContentControl":
                        ComboBoxContentControl combo = new ComboBoxContentControl(field);
                        
                        string items = "";

                        foreach (ComboBoxListItem item in combo.ListItems)
                        {
                            items += "Item: " + item.DisplayText + "\r\n";
                        }

                        MessageBox.Show("ComboBoxContentControl:\r\n" +
                            "Title: " + combo.Title + "\r\n" +
                            "Tag: " + combo.Tag + "\r\n" +
                            items);
                        break;

                    case "DateContentControl":
                        DateContentControl date = new DateContentControl(field);

                        MessageBox.Show("DateContentControl:\r\n" +
                            "Title: " + date.Title + "\r\n" +
                            "Tag: " + date.Tag + "\r\n" +
                            "Date: " + date.Date + "\r\n" +
                            "Calendar: " + date.Calendar + "\r\n" +
                            "Format: " + date.DateFormat + "\r\n");
                            
                        break;

                    case "DropDownListContentControl":
                        DropDownListContentControl drop = new DropDownListContentControl(field);

                        string dropItems = "";

                        foreach (DropDownListItem item in drop.ListItems)
                        {
                            dropItems += "Item: " + item.DisplayText + ", " + item.Value + "\r\n";
                        }

                        MessageBox.Show("DropDownListContentControl:\r\n" +
                            "Title: " + drop.Title + "\r\n" +
                            "Tag: " + drop.Tag + "\r\n" +
                            dropItems);
                        break;
                }
            }
        }
    }
}
