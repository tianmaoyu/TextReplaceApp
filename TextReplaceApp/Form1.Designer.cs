namespace TextReplaceApp
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.cob_path = new System.Windows.Forms.ComboBox();
            this.btn_close = new System.Windows.Forms.Button();
            this.btn_restore = new System.Windows.Forms.Button();
            this.btn_replace = new System.Windows.Forms.Button();
            this.btn_find = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.check_incontent = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tb_replaceText = new System.Windows.Forms.TextBox();
            this.tb_findText = new System.Windows.Forms.TextBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(1, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(134, 421);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(141, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(407, 421);
            this.tabControl1.TabIndex = 1;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.cob_path);
            this.tabPage2.Controls.Add(this.btn_close);
            this.tabPage2.Controls.Add(this.btn_restore);
            this.tabPage2.Controls.Add(this.btn_replace);
            this.tabPage2.Controls.Add(this.btn_find);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.tb_replaceText);
            this.tabPage2.Controls.Add(this.tb_findText);
            this.tabPage2.Controls.Add(this.richTextBox1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(399, 395);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "查找/替换";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // cob_path
            // 
            this.cob_path.FormattingEnabled = true;
            this.cob_path.Location = new System.Drawing.Point(97, 169);
            this.cob_path.Name = "cob_path";
            this.cob_path.Size = new System.Drawing.Size(269, 20);
            this.cob_path.TabIndex = 6;
            this.cob_path.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged_1);
            this.cob_path.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cob_path_MouseDown_1);
            // 
            // btn_close
            // 
            this.btn_close.Location = new System.Drawing.Point(315, 360);
            this.btn_close.Name = "btn_close";
            this.btn_close.Size = new System.Drawing.Size(75, 23);
            this.btn_close.TabIndex = 5;
            this.btn_close.Text = "关闭";
            this.btn_close.UseVisualStyleBackColor = true;
            this.btn_close.Click += new System.EventHandler(this.btn_close_Click);
            // 
            // btn_restore
            // 
            this.btn_restore.Location = new System.Drawing.Point(215, 360);
            this.btn_restore.Name = "btn_restore";
            this.btn_restore.Size = new System.Drawing.Size(75, 23);
            this.btn_restore.TabIndex = 5;
            this.btn_restore.Text = "还原上一版本";
            this.btn_restore.UseVisualStyleBackColor = true;
            // 
            // btn_replace
            // 
            this.btn_replace.Location = new System.Drawing.Point(116, 360);
            this.btn_replace.Name = "btn_replace";
            this.btn_replace.Size = new System.Drawing.Size(75, 23);
            this.btn_replace.TabIndex = 5;
            this.btn_replace.Text = "替换";
            this.btn_replace.UseVisualStyleBackColor = true;
            this.btn_replace.Click += new System.EventHandler(this.btn_replace_Click);
            // 
            // btn_find
            // 
            this.btn_find.Location = new System.Drawing.Point(18, 360);
            this.btn_find.Name = "btn_find";
            this.btn_find.Size = new System.Drawing.Size(75, 23);
            this.btn_find.TabIndex = 5;
            this.btn_find.Text = "查找";
            this.btn_find.UseVisualStyleBackColor = true;
            this.btn_find.Click += new System.EventHandler(this.btn_find_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.check_incontent);
            this.groupBox1.Location = new System.Drawing.Point(16, 274);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(377, 80);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "范围";
            // 
            // check_incontent
            // 
            this.check_incontent.AutoSize = true;
            this.check_incontent.Location = new System.Drawing.Point(7, 31);
            this.check_incontent.Name = "check_incontent";
            this.check_incontent.Size = new System.Drawing.Size(336, 16);
            this.check_incontent.TabIndex = 0;
            this.check_incontent.Text = "内容中查找/替换（TXT，Html,Xml,World,Excel,PPT,RTF）";
            this.check_incontent.UseVisualStyleBackColor = true;
            this.check_incontent.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 245);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "替换为：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 171);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 3;
            this.label3.Text = "选择路径：";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 208);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "查找内容：";
            // 
            // tb_replaceText
            // 
            this.tb_replaceText.Location = new System.Drawing.Point(97, 241);
            this.tb_replaceText.Name = "tb_replaceText";
            this.tb_replaceText.Size = new System.Drawing.Size(269, 21);
            this.tb_replaceText.TabIndex = 2;
            // 
            // tb_findText
            // 
            this.tb_findText.Location = new System.Drawing.Point(97, 205);
            this.tb_findText.Name = "tb_findText";
            this.tb_findText.Size = new System.Drawing.Size(269, 21);
            this.tb_findText.TabIndex = 2;
            // 
            // richTextBox1
            // 
            this.richTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBox1.Location = new System.Drawing.Point(6, 6);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(387, 141);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = resources.GetString("richTextBox1.Text");
            this.richTextBox1.TextChanged += new System.EventHandler(this.richTextBox1_TextChanged);
            // 
            // fileSystemWatcher1
            // 
            this.fileSystemWatcher1.EnableRaisingEvents = true;
            this.fileSystemWatcher1.SynchronizingObject = this;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(560, 433);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.pictureBox1);
            this.Name = "Form1";
            this.Text = "查找";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button btn_close;
        private System.Windows.Forms.Button btn_replace;
        private System.Windows.Forms.Button btn_find;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox check_incontent;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tb_replaceText;
        private System.Windows.Forms.TextBox tb_findText;
        private System.Windows.Forms.Button btn_restore;
        private System.Windows.Forms.Label label3;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.ComboBox cob_path;
    }
}

