﻿namespace Excel汇总
{
    partial class F_Fenleihuizong
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            PresentationControls.CheckBoxProperties checkBoxProperties1 = new PresentationControls.CheckBoxProperties();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(F_Fenleihuizong));
            this.btnStart = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.txtResult = new System.Windows.Forms.TextBox();
            this.txtSource = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.workCheck = new System.ComponentModel.BackgroundWorker();
            this.txtStart = new System.Windows.Forms.TextBox();
            this.txtEnd = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.txtKey = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.btnOpen = new System.Windows.Forms.Button();
            this.txtPDKeyColumn = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.rbMohu = new System.Windows.Forms.RadioButton();
            this.button5 = new System.Windows.Forms.Button();
            this.btnCheckAll = new System.Windows.Forms.Button();
            this.cbbExcuteTable = new PresentationControls.CheckBoxComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(472, 491);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(159, 32);
            this.btnStart.TabIndex = 86;
            this.btnStart.Text = "开始";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(650, 197);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 85;
            this.button3.Text = "浏览";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(650, 41);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 84;
            this.button1.Text = "浏览";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtResult
            // 
            this.txtResult.Location = new System.Drawing.Point(201, 199);
            this.txtResult.Name = "txtResult";
            this.txtResult.Size = new System.Drawing.Size(430, 21);
            this.txtResult.TabIndex = 83;
            this.txtResult.Text = "C:\\result.xls";
            // 
            // txtSource
            // 
            this.txtSource.Location = new System.Drawing.Point(201, 43);
            this.txtSource.Name = "txtSource";
            this.txtSource.Size = new System.Drawing.Size(430, 21);
            this.txtSource.TabIndex = 82;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(68, 202);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 12);
            this.label4.TabIndex = 81;
            this.label4.Text = "存放结果的文件";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(68, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(118, 12);
            this.label1.TabIndex = 80;
            this.label1.Text = "Excel文件所在目录";
            // 
            // txtStart
            // 
            this.txtStart.Location = new System.Drawing.Point(201, 147);
            this.txtStart.Name = "txtStart";
            this.txtStart.Size = new System.Drawing.Size(109, 21);
            this.txtStart.TabIndex = 92;
            this.txtStart.Text = "3";
            // 
            // txtEnd
            // 
            this.txtEnd.Location = new System.Drawing.Point(199, 173);
            this.txtEnd.Name = "txtEnd";
            this.txtEnd.Size = new System.Drawing.Size(111, 21);
            this.txtEnd.TabIndex = 91;
            this.txtEnd.Text = "1";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(68, 150);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(125, 12);
            this.label3.TabIndex = 90;
            this.label3.Text = "处理范围从第几行开始";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(68, 176);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(125, 12);
            this.label2.TabIndex = 89;
            this.label2.Text = "处理范围至倒数第几行";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(650, 491);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(159, 32);
            this.btnClose.TabIndex = 88;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // txtLog
            // 
            this.txtLog.Location = new System.Drawing.Point(70, 224);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(739, 248);
            this.txtLog.TabIndex = 87;
            // 
            // txtKey
            // 
            this.txtKey.Location = new System.Drawing.Point(201, 120);
            this.txtKey.Name = "txtKey";
            this.txtKey.Size = new System.Drawing.Size(109, 21);
            this.txtKey.TabIndex = 94;
            this.txtKey.Text = "1";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(68, 123);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(41, 12);
            this.label5.TabIndex = 93;
            this.label5.Text = "关键列";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(357, 501);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(77, 12);
            this.linkLabel1.TabIndex = 95;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "详细使用说明";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(743, 197);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(75, 23);
            this.btnOpen.TabIndex = 96;
            this.btnOpen.Text = "打开";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // txtPDKeyColumn
            // 
            this.txtPDKeyColumn.Location = new System.Drawing.Point(432, 173);
            this.txtPDKeyColumn.Name = "txtPDKeyColumn";
            this.txtPDKeyColumn.Size = new System.Drawing.Size(100, 21);
            this.txtPDKeyColumn.TabIndex = 97;
            this.txtPDKeyColumn.Text = "1";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(333, 176);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(89, 12);
            this.label6.TabIndex = 98;
            this.label6.Text = "数据判断关键列";
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(314, 100);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(71, 16);
            this.radioButton1.TabIndex = 166;
            this.radioButton1.Text = "精确匹配";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // rbMohu
            // 
            this.rbMohu.AutoSize = true;
            this.rbMohu.Checked = true;
            this.rbMohu.Location = new System.Drawing.Point(201, 100);
            this.rbMohu.Name = "rbMohu";
            this.rbMohu.Size = new System.Drawing.Size(71, 16);
            this.rbMohu.TabIndex = 165;
            this.rbMohu.TabStop = true;
            this.rbMohu.Text = "模糊匹配";
            this.rbMohu.UseVisualStyleBackColor = true;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(750, 64);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 164;
            this.button5.Text = "反选";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // btnCheckAll
            // 
            this.btnCheckAll.Location = new System.Drawing.Point(650, 64);
            this.btnCheckAll.Name = "btnCheckAll";
            this.btnCheckAll.Size = new System.Drawing.Size(75, 23);
            this.btnCheckAll.TabIndex = 163;
            this.btnCheckAll.Text = "全选";
            this.btnCheckAll.UseVisualStyleBackColor = true;
            this.btnCheckAll.Click += new System.EventHandler(this.btnCheckAll_Click);
            // 
            // cbbExcuteTable
            // 
            checkBoxProperties1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cbbExcuteTable.CheckBoxProperties = checkBoxProperties1;
            this.cbbExcuteTable.DisplayMemberSingleItem = "";
            this.cbbExcuteTable.FormattingEnabled = true;
            this.cbbExcuteTable.Location = new System.Drawing.Point(201, 66);
            this.cbbExcuteTable.Name = "cbbExcuteTable";
            this.cbbExcuteTable.Size = new System.Drawing.Size(430, 20);
            this.cbbExcuteTable.TabIndex = 162;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(68, 69);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(70, 12);
            this.label7.TabIndex = 161;
            this.label7.Text = "要处理的表";
            // 
            // F_Fenleihuizong
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(877, 565);
            this.ControlBox = false;
            this.Controls.Add(this.radioButton1);
            this.Controls.Add(this.rbMohu);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.btnCheckAll);
            this.Controls.Add(this.cbbExcuteTable);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtPDKeyColumn);
            this.Controls.Add(this.btnOpen);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.txtKey);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtResult);
            this.Controls.Add(this.txtSource);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtStart);
            this.Controls.Add(this.txtEnd);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.txtLog);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "F_Fenleihuizong";
            this.Text = "分类汇总工作表行数据";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtResult;
        private System.Windows.Forms.TextBox txtSource;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label1;
        private System.ComponentModel.BackgroundWorker workCheck;
        private System.Windows.Forms.TextBox txtStart;
        private System.Windows.Forms.TextBox txtEnd;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.TextBox txtKey;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.TextBox txtPDKeyColumn;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RadioButton rbMohu;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button btnCheckAll;
        private PresentationControls.CheckBoxComboBox cbbExcuteTable;
        private System.Windows.Forms.Label label7;
    }
}