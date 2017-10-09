using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Net;
using Rug.Osc;
using System.Diagnostics;

namespace PPAnote2osc
{
    public partial class ThisAddIn
    {
        IPAddress address;
        int port;
        OscSender oscSender;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.address = IPAddress.Parse("172.20.10.2");
            this.port = 55005;
            this.oscSender = new OscSender(address, port);
        

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            this.Application.SlideShowNextClick += Application_SlideShowNextClick;
        }

        private void Application_SlideShowNextClick(PowerPoint.SlideShowWindow Wn, PowerPoint.Effect nEffect)
        {
            //throw new NotImplementedException();
            // PowerPointスライドの「ノート」部分を取得する。
            string note = Wn.View.Slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text;
            string noteT = note.Trim();

            if (noteT != "")
            {
                oscSender.Connect();
                Debug.WriteLine("つなげた");
                oscSender.Send(new OscMessage("/scene", noteT));
                //oscSender.Send(new OscMessage("/change", int.Parse(noteT)));
                Debug.WriteLine(noteT+":");
                this.oscSender.Close();
                Debug.WriteLine("とじた");
            }
        }

        #endregion


    }
}
