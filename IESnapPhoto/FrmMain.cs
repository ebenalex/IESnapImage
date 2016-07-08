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
            
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            wbIE.ScrollBarsEnabled = false;　　//禁用滚动条
            wbIE.ScriptErrorsSuppressed = false;  //禁用脚本错误
            this.wbIE.Url = new Uri(DefaultUrl);

        }

        private void btnLogin_Click(object sender, EventArgs e)
        {


        }

        private void wbIE_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (e.Url.ToString().IndexOf("login") != -1)
            {
                wbIE.Document.GetElementById("_txtUserId").SetAttribute("value", userName);
                wbIE.Document.GetElementById("_txtPassword").SetAttribute("value", passwd);
                wbIE.Document.GetElementById("_btnLogin").InvokeMember("click");
            }
            //else
            //{
            //    btnSnap_Click(sender, e);
            //}
        }
        private string serverName = Properties.Settings.Default.ServerName;
        private string ipAddress = Properties.Settings.Default.IPAddress;
        private string DefaultUrl = Properties.Settings.Default.DefaultURL;
        private string userName = Properties.Settings.Default.UserName;
        private string passwd = Properties.Settings.Default.Password;
        private string tempDirectory = Properties.Settings.Default.TempDirectory;
        private string ServerTempDirectory = Properties.Settings.Default.ServerTempDirectory;
        private string connectionString = Properties.Settings.Default.ConnectionString;
        private void timGetFormPicture_Tick(object sender, EventArgs e)
        {
            
            timGetFormPicture.Enabled = false;

            //连接数据库后将截图插入至附件字段
            SqlConnection sqlconn = new SqlConnection();
            sqlconn.ConnectionString = connectionString;
            SqlConnection sqlconn1 = new SqlConnection();
            sqlconn1.ConnectionString = sqlconn.ConnectionString;
            sqlconn1.Open();

            #region Variable 
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
            string[] powerUsers = Properties.Settings.Default.PowserUser.Split(';');
            //限制流程列表
            string[] limitedProcess = Properties.Settings.Default.ProcessLimited.Split(';');



            string taskID = "0";
            string ftaskKey = "tid=";
            string etaskKey = "&";
            string processStepID = "0";
            string pStepKey = "&pid";

            string commandStr = "";
            string ownerAccount = "";

            string mailKey = "邮件审批：";
            string tmpMessage = "";

            #endregion

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
                    #region 根据邮件的内容进行分析读取相关变量的内容
                    //获取审批流水号
                    title = reader["Title"].ToString();
                    lsCode = title.Substring(title.LastIndexOf('：')+1, title.Length- title.LastIndexOf('：')-1);
                    if (title.IndexOf("新任务") > -1)
                    {
                        //获取快速审批路径
                        message = reader["Message"].ToString();
                        startPos = message.IndexOf(fKeyString);
                        startPos += fKeyString.Length;

                        endPos = message.IndexOf(eKeyString, startPos);

                        processURL = message.Substring(startPos, endPos - startPos);
                        processURL = processURL.Replace(".113", ipAddress);
                        startPos = processURL.IndexOf(ftaskKey) + ftaskKey.Length;
                        endPos = processURL.IndexOf(etaskKey, startPos);
                        taskID = processURL.Substring(startPos, endPos - startPos);

                        startPos = processURL.IndexOf(pStepKey) + pStepKey.Length;
                        endPos = processURL.Length;
                        processStepID = processURL.Substring(startPos + 1, endPos - startPos - 1);
                        //读取邮件审批，若不是审批人员，则将此功能取消
                        startPos = message.IndexOf(mailKey);
                        if (startPos != -1)
                        {
                            tmpMessage = message.Substring(0, startPos - 1);
                        }
                        else
                        {
                            tmpMessage = message;
                        }

                        if (!powerUsers.Contains(ownerAccount.ToLower()))
                        {
                            //若不是高级领导则删除邮件审批的内容
                            commandStr = "Udpate set [Message] = case when CHARINDEX ('邮件审批',[Message] ,0)>0 then substring([Message],0,CHARINDEX ('邮件审批',[Message] ,0)) else [message] end  where MessageID= " + reader["MessageID"].ToString();
                            SqlCommand cmd4 = new SqlCommand(commandStr, sqlconn1);
                            cmd.ExecuteNonQuery();
                        }

                        #endregion

                        if (!limitedProcess.Contains(lsCode.Substring(0, 4)))
                        {
                            #region 读取用户名称

                            commandStr = "Select OwnerAccount From BPMInstProcSteps where StepID = " + processStepID + " and TaskID = " + taskID;
                            SqlCommand cmd2 = new SqlCommand(commandStr, sqlconn1);
                            object oOwnerAccount = cmd2.ExecuteScalar();
                            ownerAccount = oOwnerAccount.ToString();
                            #endregion

                            #region 生成文件并同时发送文件
                            if (powerUsers.Contains(ownerAccount.ToLower()))
                            {

                                //更新审批人为后台运行帐号
                                commandStr = "Update BPMInstProcSteps set OwnerAccount='" + userName + "' where StepID = " + processStepID + " and TaskID = " + taskID;
                                cmd2 = new SqlCommand(commandStr, sqlconn1);
                                cmd2.ExecuteNonQuery();
                                //获取附件的名称
                                Bitmap bit = EVWB.GetHtmlToImage.GetHtmlImage(new Uri(processURL), 1200);
                                EVFile.FileInfo fi = new EVFile.FileInfo(tempDirectory, "jpg");
                                bit.Save(fi.Filepath, System.Drawing.Imaging.ImageFormat.Jpeg);
                                if (reader["Attachments"] is DBNull)
                                {
                                    attachment = "Process.jpg=" + fi.Filepath;
                                }
                                else
                                {
                                    attachment = reader["Attachments"].ToString() + ";" + "Process.jpg=" + fi.Filepath;
                                }

                                //更新附件的状态
                                commandStr = "Update BPMSysMessagesQueuePrepare Set Attachments='" + attachment + "' where MessageID= " + reader["MessageID"].ToString();

                                cmd2 = new SqlCommand(commandStr, sqlconn1);
                                cmd2.ExecuteNonQuery();
                                //还原原始的帐户
                                commandStr = "Update BPMInstProcSteps set OwnerAccount='" + ownerAccount + "' where StepID = " + processStepID + " and TaskID = " + taskID;
                                cmd2 = new SqlCommand(commandStr, sqlconn1);
                                cmd2.ExecuteNonQuery();

                                //拷贝文件至指定的目录
                                //创建指定的临时目录
                                //if (!System.IO.Directory.Exists(ServerTempDirectory))
                                //{
                                //    System.IO.Directory.CreateDirectory(ServerTempDirectory);
                                //}
                                string copyString = "copy " + fi.Filepath + " " + ServerTempDirectory + fi.Filename + " /Y";
                                System.Diagnostics.Process p = new System.Diagnostics.Process();
                                p.StartInfo = new System.Diagnostics.ProcessStartInfo("cmd.exe", " /C " + copyString);
                                p.Start();

                            }
                            #endregion
                        }
                    }

                    #region 恢复邮件列表及删除准备的资料
                    //将准备好数据插入至发邮件的队列
                    commandStr = "insert into BPMSysMessagesQueue([ProviderName],[Address],[Title] ,[Message] ,[CreateAt],[LastSendAt] ,[FailCount],[Attachments])  Select '"+ provider+"',[Address] ,[Title] ,[Message] ,[CreateAt] ,[LastSendAt],[FailCount] ,[Attachments] from BPMSysMessagesQueuePrepare where MessageID= " + reader["MessageID"].ToString() ;

                    SqlCommand cmd1 = new SqlCommand(commandStr, sqlconn1);
                    cmd1.ExecuteNonQuery();

                    commandStr = "delete from BPMSysMessagesQueuePrepare where MessageID= " + reader["MessageID"].ToString();

                    SqlCommand cmd3 = new SqlCommand(commandStr, sqlconn1);
                    cmd3.ExecuteNonQuery();

                    #endregion
                    
                }

               
            }

            sqlconn1.Close();
            this.wbIE.Url = new Uri(DefaultUrl);
            Task.Delay(10000);

            timGetFormPicture.Enabled = true;
            
        }

    
    }
}
