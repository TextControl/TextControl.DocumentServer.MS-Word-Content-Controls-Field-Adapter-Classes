using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TXTextControl.DocumentServer.Fields
{
    public class PlainTextContentControl : ContentControlFieldAdapter
    {
        /*-------------------------------------------------------------------------------------------------------
        ** Constructors
        **-----------------------------------------------------------------------------------------------------*/

        public PlainTextContentControl(ApplicationField ApplicationField)
            : base(ApplicationField)
        {

        }

        public PlainTextContentControl()
            : base()
        {

        }

        /*-------------------------------------------------------------------------------------------------------
        ** Properties
        **-----------------------------------------------------------------------------------------------------*/

        public string Text
        {
            get { return this.ApplicationField.Text; }
            set
            {
                this.ApplicationField.Text = value;
            }
        }

        /*-------------------------------------------------------------------------------------------------------
        ** Methods
        **-----------------------------------------------------------------------------------------------------*/

        protected override void SetParameters()
        {
            this.ApplicationField.Parameters =
                new string[] { this.GetXmlBaseStructure().OuterXml };
        }
    }
}
