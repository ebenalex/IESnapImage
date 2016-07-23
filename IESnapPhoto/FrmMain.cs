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
using System.Diagnostics;
using System.IO;

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

            try
            {

                //连接数据库后将截图插入至附件字段
                SqlConnection sqlconn = new SqlConnection();
                sqlconn.ConnectionString = connectionString;

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

                string ownerAccount = "";
                #endregion

                //循环处理时，当处理完成一条，则将此条写入邮件序列中
                //限制使用的帐户列表以及限制不使用此流程的列表
                using (sqlconn)
                {
                    sqlconn.Open();

                    //遍历数据库发邮件的准备表
                    SqlCommand cmd = new SqlCommand("SELECT * FROM BPMSysMessagesQueuePrepare", sqlconn);
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    //SqlDataReader reader = cmd.ExecuteReader();
                    foreach (DataRow reader in ds.Tables[0].Rows)
                    {

                        #region 根据邮件的内容进行分析读取相关变量的内容
                        //获取审批流水号
                        title = reader["Title"].ToString();
                        lsCode = title.Substring(title.LastIndexOf('：') + 1, title.Length - title.LastIndexOf('：') - 1);
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

                            ownerAccount = readOwnerAccount(sqlconn, taskID, processStepID);

                            processMesageEmailReply(sqlconn, powerUsers, ownerAccount, reader["MessageID"].ToString(), message);

                            #endregion

                            if (!limitedProcess.Contains(lsCode.Substring(0, 4)))
                            {


                                #region 生成文件并同时发送文件
                                if (powerUsers.Contains(ownerAccount.ToLower()))
                                {
                                    EVFile.FileInfo fi;
                                    fi = snapPictureSendMail(sqlconn, processURL, taskID, processStepID, ownerAccount, reader);

                                    copyFiles(fi);

                                }
                                #endregion
                            }
                        }

                        updateSendMail(sqlconn, provider, reader);
                    }


                }


            }
            catch (Exception ee)
            {
                writeSystemEventErr(ee);

            }
            finally
            {
                this.wbIE.Url = new Uri(DefaultUrl);
                Task.Delay(10000);

                timGetFormPicture.Enabled = true;
            }

        }

        /// <summary>
        /// 更新数据库并发送邮件
        /// </summary>
        /// <param name="sqlconn"> DB connection</param>
        /// <param name="provider"> Mail provider</param>
        /// <param name="reader"></param>
        private static void updateSendMail(SqlConnection sqlconn, string provider, DataRow reader)
        {
            #region 恢复邮件列表及删除准备的资料
            //将准备好数据插入至发邮件的队列
            string commandStr = "insert into BPMSysMessagesQueue([ProviderName],[Address],[Title] ,[Message] ,[CreateAt],[LastSendAt] ,[FailCount],[Attachments])  Select '" + provider + "',[Address] ,[Title] ,[Message] ,[CreateAt] ,[LastSendAt],[FailCount] ,[Attachments] from BPMSysMessagesQueuePrepare where MessageID= " + reader["MessageID"].ToString();

            SqlCommand cmd1 = new SqlCommand(commandStr, sqlconn);
            cmd1.ExecuteNonQuery();

            commandStr = "delete from BPMSysMessagesQueuePrepare where MessageID= " + reader["MessageID"].ToString();

            SqlCommand cmd3 = new SqlCommand(commandStr, sqlconn);
            cmd3.ExecuteNonQuery();

            #endregion
        }

        /// <summary>
        /// 生成图片并更新附件中文件信息
        /// </summary>
        /// <param name="sqlconn">DB Connection</param>
        /// <param name="processURL">Process URL </param>
        /// <param name="taskID">Task ID</param>
        /// <param name="processStepID">Process Step ID</param>
        /// <param name="ownerAccount">My Account</param>
        /// <param name="reader"></param>
        /// <returns></returns>
        private EVFile.FileInfo snapPictureSendMail(SqlConnection sqlconn, string processURL, string taskID, string processStepID, string ownerAccount, DataRow reader)
        {
            //更新审批人为后台运行帐号

            string commandStr = "Update BPMInstProcSteps set OwnerAccount='" + userName + "' where StepID = " + processStepID + " and TaskID = " + taskID;
            SqlCommand cmd2 = new SqlCommand(commandStr, sqlconn);
            cmd2.ExecuteNonQuery();
            //获取附件的名称
            Bitmap bit = EVWB.GetHtmlToImage.GetHtmlImage(new Uri(processURL), 1200);
            EVFile.FileInfo fi = new EVFile.FileInfo(tempDirectory, "jpg");
            bit.Save(fi.Filepath, System.Drawing.Imaging.ImageFormat.Jpeg);
            string attachment = "";
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

            cmd2 = new SqlCommand(commandStr, sqlconn);
            cmd2.ExecuteNonQuery();
            //还原原始的帐户
            commandStr = "Update BPMInstProcSteps set OwnerAccount='" + ownerAccount + "' where StepID = " + processStepID + " and TaskID = " + taskID;
            cmd2 = new SqlCommand(commandStr, sqlconn);
            cmd2.ExecuteNonQuery();
            return fi;
        }

        /// <summary>
        /// 将临时生成的文件拷贝至目的文件夹
        /// </summary>
        /// <param name="fi">文件</param>
        private void copyFiles(EVFile.FileInfo fi)
        {
            string copyString = "copy " + fi.Filepath + " " + ServerTempDirectory + fi.Filename + " /Y";
            if(System.IO.File.Exists(fi.Filepath))
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo = new System.Diagnostics.ProcessStartInfo("cmd.exe", " /C " + copyString);
                p.Start();
            }
            
        }

        private static string readOwnerAccount(SqlConnection sqlconn, string taskID, string processStepID)
        {
            #region 读取用户名称
            string commandStr = "Select OwnerAccount From BPMInstProcSteps where StepID = " + processStepID + " and TaskID = " + taskID;
            SqlCommand cmd2 = new SqlCommand(commandStr, sqlconn);
            object oOwnerAccount = cmd2.ExecuteScalar();
            string ownerAccount = oOwnerAccount.ToString();
            return ownerAccount;
            #endregion
        }

        /// <summary>
        /// 处理用户是否为高级用户，如果不为高级用户，则将邮件回复审批的功能取消
        /// </summary>
        /// <param name="sqlconn">DB Conneciton</param>
        /// <param name="powerUsers">Power User List</param>
        /// <param name="ownerAccount">My User Account</param>
        /// <param name="reader"></param>
        private static void processMesageEmailReply(SqlConnection sqlconn, string[] powerUsers, string ownerAccount, string messageID,string message)
        {
            string mailKey = "邮件审批：";
            if (!powerUsers.Contains(ownerAccount.ToLower()) && message.IndexOf(mailKey)> -1)
            {
                //若不是高级领导则删除邮件审批的内容
                string commandStr = "Update  BPMSysMessagesQueuePrepare  set [Message] = case when CHARINDEX ('" + mailKey + "',[Message] ,0)>0 then substring([Message],0,CHARINDEX ('" + mailKey + "',[Message] ,0)) else [message] end  where MessageID= " + messageID;
                SqlCommand cmd4 = new SqlCommand(commandStr, sqlconn);
                cmd4.ExecuteNonQuery();
            }
        }

        private static void writeSystemEventErr(Exception ee)
        {
            StreamWriter sw  = File.AppendText("Log.txt");
            sw.WriteLine(System.DateTime.Now.ToString() + " : " + ee.ToString());
            sw.Close();
            
            //if (!EventLog.SourceExists("IESnapPhoto"))
            //{
            //    EventLog.CreateEventSource("IESnapPhoto", "BPMSnapPhoto");

            //}
            //EventLog myLog = new EventLog();
            //myLog.Source = "IESnapPhoto";
            //myLog.WriteEntry(ee.ToString(), EventLogEntryType.Error, 1, 0);
        }

    }
}
