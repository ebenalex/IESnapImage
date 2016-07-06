using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;


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

        private void timGetFormPicture_Tick(object sender, EventArgs e)
        {
            timGetFormPicture.Stop();
            timGetFormPicture.Enabled = false;

            //连接数据库后将截图插入至附件字段
            SqlConnection sqlconn = new SqlConnection();
            sqlconn.ConnectionString = @"Server=RCA064;Database=BPMDB;User ID=sa;Password=Call3248;Trusted_Connection=False";
            SqlConnection sqlconn1 = new SqlConnection();
            sqlconn1.ConnectionString = @"Server=RCA064;Database=BPMDB;User ID=sa;Password=Call3248;Trusted_Connection=False";
            sqlconn1.Open();
            //审批流水号
            string lsCode = "";
            string title = "";
            string message = "";
            //处理的快速路径
            string processURL = "";
            string fKeyString = @"快速处理：<a href=""";
            string eKeyString = @""" target=""_blank"">";

            int startPos = 0;
            int endPos = 0;
            //附件的内容
            string attachment = "";
            
            //邮件的处理方式
            string provider = "EVLMail";

            //设置用户列表
            string[] powerUsers = new string[] { "dschang" };
            //限制流程列表
            string[] limitedProcess = new string[] { "ITIN" };

            string userAccount = "";
            string userKey = "@";


            string taskID = "0";
            string ftaskKey = "tid=";
            string etaskKey = "&";
            string processStepID = "0";
            string pStepKey = "&pid";

            string commandStr = "";
            string ownerAccount = "";

            //循环处理时，当处理完成一条，则将此条写入邮件序列中
            //限制使用的帐户列表以及限制不使用此流程的列表
            using (sqlconn)
            {
                sqlconn.Open();
               
                //遍历数据库发邮件的准备表
                SqlCommand cmd = new SqlCommand("SELECT * FROM BPMSysMessagesQueuePrepare", sqlconn);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    //获取审批流水号
                    title = reader["Title"].ToString();
                    lsCode = title.Substring(title.LastIndexOf('：')+1, title.Length- title.LastIndexOf('：')-1);

                    //获取快速审批路径
                    message = reader["Message"].ToString();
                    startPos = message.IndexOf(fKeyString);
                    startPos += fKeyString.Length;

                    endPos = message.IndexOf(eKeyString, startPos);

                    processURL = message.Substring(startPos, endPos - startPos);
                    processURL = processURL.Replace(".113", ".64");
                    startPos = processURL.IndexOf(ftaskKey)+ftaskKey.Length;
                    endPos = processURL.IndexOf(etaskKey,startPos);
                    taskID = processURL.Substring(startPos, endPos - startPos );

                    startPos = processURL.IndexOf(pStepKey)+pStepKey.Length;
                    endPos = processURL.Length;
                    processStepID = processURL.Substring(startPos + 1, endPos - startPos - 1);

                    if (!limitedProcess.Contains(lsCode.Substring(0, 4)))
                    {
                        endPos = reader["Address"].ToString().IndexOf(userKey);
                        userAccount = reader["Address"].ToString().Substring(0, endPos);
                        if (powerUsers.Contains(userAccount))
                        {

                            //更新审批人为后台运行帐号
                            
                            commandStr = "Select OwnerAccount From BPMInstProcSteps where StepID = " + processStepID + " and TaskID = " + taskID;
                            SqlCommand cmd2 = new SqlCommand(commandStr, sqlconn1);
                            object oOwnerAccount = cmd2.ExecuteScalar();
                            ownerAccount = oOwnerAccount.ToString();
                            commandStr = "Update BPMInstProcSteps set OwnerAccount='dschang' where StepID = " + processStepID + " and TaskID = " + taskID;
                            cmd2 = new SqlCommand(commandStr, sqlconn1);
                            cmd2.ExecuteNonQuery();
                            //获取附件的名称
                            Bitmap bit = EVWB.GetHtmlToImage.GetHtmlImage(new Uri(processURL), 1200);
                            EVFile.FileInfo fi = new EVFile.FileInfo("d:\\", "jpg");
                            bit.Save(fi.Filepath, System.Drawing.Imaging.ImageFormat.Jpeg);
                            if (reader["Attachments"] is DBNull)
                            {
                                attachment = "Process.jpg=" + fi.Filepath;
                            }
                            else
                            {
                                attachment = reader["Attachments"].ToString() + ";" + "Process.jpg=" + fi.Filepath;
                            }
                            //还原原始的帐户
                            commandStr = "Update BPMInstProcSteps set OwnerAccount='" + ownerAccount + "' where StepID = " + processStepID + " and TaskID = " + taskID;
                            cmd2 = new SqlCommand(commandStr, sqlconn1);
                            cmd2.ExecuteNonQuery();
                        }
                    }

                    //将准备好数据插入至发邮件的队列
                    commandStr = "insert into BPMSysMessagesQueue([ProviderName],[Address],[Title] ,[Message] ,[CreateAt],[LastSendAt] ,[FailCount],[Attachments])  Select '"+ provider+"',[Address] ,[Title] ,[Message] ,[CreateAt] ,[LastSendAt],[FailCount] ,Attachments] from BPMSysMessagesQueuePrepare where MessageID= " + reader["MessageID"].ToString() ;

                    SqlCommand cmd1 = new SqlCommand(commandStr, sqlconn1);
                    cmd1.ExecuteNonQuery();

                    commandStr = "delete from BPMSysMessagesQueuePrepare where MessageID= " + reader["MessageID"].ToString();

                    SqlCommand cmd3 = new SqlCommand(commandStr, sqlconn1);
                    cmd3.ExecuteNonQuery();





                }
            }

            sqlconn1.Close();

            timGetFormPicture.Enabled = true;
            timGetFormPicture.Start();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            timGetFormPicture_Tick(sender, e);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Bitmap bit = EVWB.GetHtmlToImage.GetHtmlImage(new Uri(@"http://10.151.65.64/BPM/YZSoft/Forms/XForm/ITForms/ITInResRequest50.aspx?tid=3273"), 1200);
            EVFile.FileInfo fi = new EVFile.FileInfo("d:\\", "jpg");
            bit.Save(fi.Filepath, System.Drawing.Imaging.ImageFormat.Jpeg);
            
        }
    }
}
