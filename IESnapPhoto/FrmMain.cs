using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IESnapPhoto
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }
        private void btnSnap_Click(object sender, EventArgs e)
        {
            Bitmap bit = EVWB.GetHtmlToImage.GetHtmlImage(new Uri(@"http://10.151.65.113/BPM/YZSoft/Forms/XForm/ExpForms/ExpCenForms/Exp_Cen_GeneralCost2.aspx?app=ExpFormService%2fExpCenFormService%2fExp_Cen_GeneralCost&token=dae18004a26d1c6ef735703503f83edb6f03e02e&tid=6499&key=6499&md=App&formstate=Read"), 1200);
            EVFile.FileInfo fi = new EVFile.FileInfo("d:\\", "jpg");
            bit.Save(fi.Filepath, System.Drawing.Imaging.ImageFormat.Jpeg);
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            wbIE.ScrollBarsEnabled = false;　　//禁用滚动条
            wbIE.ScriptErrorsSuppressed = false;  //禁用脚本错误
           
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {


        }

        private void wbIE_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (e.Url.ToString().IndexOf("login") != -1)
            {
                wbIE.Document.GetElementById("_txtUserId").SetAttribute("value", "dschang");
                wbIE.Document.GetElementById("_txtPassword").SetAttribute("value", "");
                wbIE.Document.GetElementById("_btnLogin").InvokeMember("click");
            }
            else
            {
                btnSnap_Click(sender, e);
            }
        }
    }
}
