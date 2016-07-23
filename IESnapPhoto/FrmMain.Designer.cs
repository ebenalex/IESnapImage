namespace IESnapPhoto
{
    partial class FrmMain
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.wbIE = new System.Windows.Forms.WebBrowser();
            this.timGetFormPicture = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // wbIE
            // 
            this.wbIE.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.wbIE.Location = new System.Drawing.Point(3, 3);
            this.wbIE.MinimumSize = new System.Drawing.Size(20, 20);
            this.wbIE.Name = "wbIE";
            this.wbIE.Size = new System.Drawing.Size(525, 435);
            this.wbIE.TabIndex = 0;
            this.wbIE.Url = new System.Uri("", System.UriKind.Relative);
            this.wbIE.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.wbIE_DocumentCompleted);
            // 
            // timGetFormPicture
            // 
            this.timGetFormPicture.Enabled = true;
            this.timGetFormPicture.Interval = 10000;
            this.timGetFormPicture.Tick += new System.EventHandler(this.timGetFormPicture_Tick);
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(529, 440);
            this.Controls.Add(this.wbIE);
            this.Name = "FrmMain";
            this.Text = "邮件中加入审批页面";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.WebBrowser wbIE;
        private System.Windows.Forms.Timer timGetFormPicture;
    }
}

