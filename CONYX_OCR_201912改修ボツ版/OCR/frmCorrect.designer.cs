namespace CONYX_OCR.OCR
{
    partial class frmCorrect
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCorrect));
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMonth = new System.Windows.Forms.TextBox();
            this.txtYear = new System.Windows.Forms.TextBox();
            this.hScrollBar1 = new System.Windows.Forms.HScrollBar();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btnMinus = new System.Windows.Forms.Button();
            this.btnPlus = new System.Windows.Forms.Button();
            this.btnEnd = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnBefore = new System.Windows.Forms.Button();
            this.btnFirst = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnDataMake = new System.Windows.Forms.Button();
            this.btnRtn = new System.Windows.Forms.Button();
            this.btnErrCheck = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.lblNoImage = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.leadImg = new Leadtools.WinForms.RasterImageViewer();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblErrMsg = new System.Windows.Forms.Label();
            this.lblName = new System.Windows.Forms.Label();
            this.txtCode = new System.Windows.Forms.TextBox();
            this.lblPage = new System.Windows.Forms.Label();
            this.chkKakunin = new System.Windows.Forms.CheckBox();
            this.txtMemo = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.dGV = new CONYX_OCR.DataGridViewEx();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dGV)).BeginInit();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("游ゴシック", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(812, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(17, 13);
            this.label3.TabIndex = 56;
            this.label3.Text = "月";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("游ゴシック", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(739, 78);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(17, 13);
            this.label2.TabIndex = 55;
            this.label2.Text = "年";
            // 
            // txtMonth
            // 
            this.txtMonth.BackColor = System.Drawing.Color.White;
            this.txtMonth.Font = new System.Drawing.Font("游ゴシック", 29.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMonth.ForeColor = System.Drawing.Color.DarkBlue;
            this.txtMonth.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtMonth.Location = new System.Drawing.Point(756, 48);
            this.txtMonth.MaxLength = 2;
            this.txtMonth.Name = "txtMonth";
            this.txtMonth.Size = new System.Drawing.Size(56, 70);
            this.txtMonth.TabIndex = 1;
            this.txtMonth.Text = "12";
            this.txtMonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtMonth.WordWrap = false;
            this.txtMonth.TextChanged += new System.EventHandler(this.txtYear_TextChanged);
            this.txtMonth.Enter += new System.EventHandler(this.txtYear_Enter);
            this.txtMonth.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            this.txtMonth.Leave += new System.EventHandler(this.txtYear_Leave);
            // 
            // txtYear
            // 
            this.txtYear.BackColor = System.Drawing.Color.White;
            this.txtYear.Font = new System.Drawing.Font("游ゴシック", 29.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtYear.ForeColor = System.Drawing.Color.DarkBlue;
            this.txtYear.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtYear.Location = new System.Drawing.Point(638, 48);
            this.txtYear.MaxLength = 4;
            this.txtYear.Name = "txtYear";
            this.txtYear.Size = new System.Drawing.Size(101, 70);
            this.txtYear.TabIndex = 0;
            this.txtYear.Text = "2018";
            this.txtYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtYear.WordWrap = false;
            this.txtYear.TextChanged += new System.EventHandler(this.txtYear_TextChanged);
            this.txtYear.Enter += new System.EventHandler(this.txtYear_Enter);
            this.txtYear.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            this.txtYear.Leave += new System.EventHandler(this.txtYear_Leave);
            // 
            // hScrollBar1
            // 
            this.hScrollBar1.Location = new System.Drawing.Point(234, 903);
            this.hScrollBar1.Name = "hScrollBar1";
            this.hScrollBar1.Size = new System.Drawing.Size(392, 34);
            this.hScrollBar1.TabIndex = 15;
            this.toolTip1.SetToolTip(this.hScrollBar1, "出勤簿を移動します");
            this.hScrollBar1.Scroll += new System.Windows.Forms.ScrollEventHandler(this.hScrollBar1_Scroll);
            // 
            // toolTip1
            // 
            this.toolTip1.BackColor = System.Drawing.Color.LemonChiffon;
            // 
            // btnMinus
            // 
            this.btnMinus.BackColor = System.Drawing.SystemColors.Control;
            this.btnMinus.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMinus.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnMinus.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnMinus.Image = ((System.Drawing.Image)(resources.GetObject("btnMinus.Image")));
            this.btnMinus.Location = new System.Drawing.Point(45, 903);
            this.btnMinus.Name = "btnMinus";
            this.btnMinus.Size = new System.Drawing.Size(37, 34);
            this.btnMinus.TabIndex = 10;
            this.btnMinus.TabStop = false;
            this.btnMinus.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolTip1.SetToolTip(this.btnMinus, "画像を縮小表示します");
            this.btnMinus.UseVisualStyleBackColor = false;
            this.btnMinus.Click += new System.EventHandler(this.btnMinus_Click);
            // 
            // btnPlus
            // 
            this.btnPlus.BackColor = System.Drawing.SystemColors.Control;
            this.btnPlus.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPlus.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnPlus.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnPlus.Image = ((System.Drawing.Image)(resources.GetObject("btnPlus.Image")));
            this.btnPlus.Location = new System.Drawing.Point(9, 903);
            this.btnPlus.Name = "btnPlus";
            this.btnPlus.Size = new System.Drawing.Size(37, 34);
            this.btnPlus.TabIndex = 9;
            this.btnPlus.TabStop = false;
            this.btnPlus.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolTip1.SetToolTip(this.btnPlus, "画像を拡大表示します");
            this.btnPlus.UseVisualStyleBackColor = false;
            this.btnPlus.Click += new System.EventHandler(this.btnPlus_Click);
            // 
            // btnEnd
            // 
            this.btnEnd.BackColor = System.Drawing.SystemColors.Control;
            this.btnEnd.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEnd.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnEnd.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnEnd.Image = ((System.Drawing.Image)(resources.GetObject("btnEnd.Image")));
            this.btnEnd.Location = new System.Drawing.Point(189, 903);
            this.btnEnd.Name = "btnEnd";
            this.btnEnd.Size = new System.Drawing.Size(37, 34);
            this.btnEnd.TabIndex = 14;
            this.btnEnd.TabStop = false;
            this.toolTip1.SetToolTip(this.btnEnd, "最後尾の出勤簿データへ移動します");
            this.btnEnd.UseVisualStyleBackColor = false;
            this.btnEnd.Click += new System.EventHandler(this.btnEnd_Click);
            // 
            // btnNext
            // 
            this.btnNext.BackColor = System.Drawing.SystemColors.Control;
            this.btnNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNext.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnNext.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnNext.Image = ((System.Drawing.Image)(resources.GetObject("btnNext.Image")));
            this.btnNext.Location = new System.Drawing.Point(153, 903);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(37, 34);
            this.btnNext.TabIndex = 13;
            this.btnNext.TabStop = false;
            this.toolTip1.SetToolTip(this.btnNext, "次の出勤簿データへ移動します");
            this.btnNext.UseVisualStyleBackColor = false;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnBefore
            // 
            this.btnBefore.BackColor = System.Drawing.SystemColors.Control;
            this.btnBefore.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBefore.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnBefore.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnBefore.Image = ((System.Drawing.Image)(resources.GetObject("btnBefore.Image")));
            this.btnBefore.Location = new System.Drawing.Point(117, 903);
            this.btnBefore.Name = "btnBefore";
            this.btnBefore.Size = new System.Drawing.Size(37, 34);
            this.btnBefore.TabIndex = 12;
            this.btnBefore.TabStop = false;
            this.toolTip1.SetToolTip(this.btnBefore, "前の出勤簿データへ移動します");
            this.btnBefore.UseVisualStyleBackColor = false;
            this.btnBefore.Click += new System.EventHandler(this.btnBefore_Click);
            // 
            // btnFirst
            // 
            this.btnFirst.BackColor = System.Drawing.SystemColors.Control;
            this.btnFirst.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFirst.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnFirst.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnFirst.Image = ((System.Drawing.Image)(resources.GetObject("btnFirst.Image")));
            this.btnFirst.Location = new System.Drawing.Point(81, 903);
            this.btnFirst.Name = "btnFirst";
            this.btnFirst.Size = new System.Drawing.Size(37, 34);
            this.btnFirst.TabIndex = 11;
            this.btnFirst.TabStop = false;
            this.toolTip1.SetToolTip(this.btnFirst, "先頭の出勤簿データへ移動します");
            this.btnFirst.UseVisualStyleBackColor = false;
            this.btnFirst.Click += new System.EventHandler(this.btnFirst_Click);
            // 
            // btnDel
            // 
            this.btnDel.BackColor = System.Drawing.SystemColors.Control;
            this.btnDel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDel.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDel.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnDel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnDel.Location = new System.Drawing.Point(901, 903);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(111, 34);
            this.btnDel.TabIndex = 8;
            this.btnDel.Text = "出勤簿削除";
            this.toolTip1.SetToolTip(this.btnDel, "表示中の出勤簿データを削除します");
            this.btnDel.UseVisualStyleBackColor = false;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnDataMake
            // 
            this.btnDataMake.BackColor = System.Drawing.SystemColors.Control;
            this.btnDataMake.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDataMake.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDataMake.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnDataMake.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnDataMake.Location = new System.Drawing.Point(771, 903);
            this.btnDataMake.Name = "btnDataMake";
            this.btnDataMake.Size = new System.Drawing.Size(127, 34);
            this.btnDataMake.TabIndex = 7;
            this.btnDataMake.Text = "汎用データ作成";
            this.toolTip1.SetToolTip(this.btnDataMake, "エラーチェックの後、勤怠データを作成します");
            this.btnDataMake.UseVisualStyleBackColor = false;
            this.btnDataMake.Click += new System.EventHandler(this.btnDataMake_Click);
            // 
            // btnRtn
            // 
            this.btnRtn.BackColor = System.Drawing.SystemColors.Control;
            this.btnRtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRtn.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRtn.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnRtn.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRtn.Location = new System.Drawing.Point(1129, 903);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(94, 34);
            this.btnRtn.TabIndex = 10;
            this.btnRtn.Text = "終了";
            this.toolTip1.SetToolTip(this.btnRtn, "プログラムを終了しメニューへ戻ります");
            this.btnRtn.UseVisualStyleBackColor = false;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // btnErrCheck
            // 
            this.btnErrCheck.BackColor = System.Drawing.SystemColors.Control;
            this.btnErrCheck.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnErrCheck.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnErrCheck.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnErrCheck.Location = new System.Drawing.Point(638, 903);
            this.btnErrCheck.Name = "btnErrCheck";
            this.btnErrCheck.Size = new System.Drawing.Size(130, 34);
            this.btnErrCheck.TabIndex = 6;
            this.btnErrCheck.Text = "エラーチェック";
            this.toolTip1.SetToolTip(this.btnErrCheck, "エラーチェックを実行します");
            this.btnErrCheck.UseVisualStyleBackColor = false;
            this.btnErrCheck.Click += new System.EventHandler(this.btnErrCheck_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.Control;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(1015, 903);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(111, 34);
            this.button1.TabIndex = 9;
            this.button1.Text = "画像印刷";
            this.toolTip1.SetToolTip(this.button1, "表示中の出勤簿データを削除します");
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblNoImage
            // 
            this.lblNoImage.BackColor = System.Drawing.SystemColors.Control;
            this.lblNoImage.Font = new System.Drawing.Font("游ゴシック", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNoImage.ForeColor = System.Drawing.Color.White;
            this.lblNoImage.Location = new System.Drawing.Point(154, 393);
            this.lblNoImage.Name = "lblNoImage";
            this.lblNoImage.Size = new System.Drawing.Size(322, 42);
            this.lblNoImage.TabIndex = 119;
            this.lblNoImage.Text = "画像はありません";
            this.lblNoImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(9, 8);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(619, 883);
            this.pictureBox1.TabIndex = 120;
            this.pictureBox1.TabStop = false;
            // 
            // leadImg
            // 
            this.leadImg.AutoScroll = false;
            this.leadImg.BackColor = System.Drawing.SystemColors.Control;
            this.leadImg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.leadImg.Location = new System.Drawing.Point(9, 8);
            this.leadImg.Name = "leadImg";
            this.leadImg.Size = new System.Drawing.Size(617, 889);
            this.leadImg.TabIndex = 121;
            this.leadImg.MouseLeave += new System.EventHandler(this.leadImg_MouseLeave);
            this.leadImg.MouseMove += new System.Windows.Forms.MouseEventHandler(this.leadImg_MouseMove);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.lblErrMsg);
            this.panel1.Location = new System.Drawing.Point(638, 7);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(588, 38);
            this.panel1.TabIndex = 162;
            // 
            // lblErrMsg
            // 
            this.lblErrMsg.BackColor = System.Drawing.Color.White;
            this.lblErrMsg.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblErrMsg.Font = new System.Drawing.Font("ＭＳ ゴシック", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblErrMsg.ForeColor = System.Drawing.Color.Red;
            this.lblErrMsg.Location = new System.Drawing.Point(0, 0);
            this.lblErrMsg.Name = "lblErrMsg";
            this.lblErrMsg.Size = new System.Drawing.Size(584, 34);
            this.lblErrMsg.TabIndex = 0;
            this.lblErrMsg.Text = "label33";
            this.lblErrMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblName
            // 
            this.lblName.BackColor = System.Drawing.Color.White;
            this.lblName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblName.Font = new System.Drawing.Font("游ゴシック", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblName.Location = new System.Drawing.Point(958, 61);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(268, 36);
            this.lblName.TabIndex = 270;
            this.lblName.Text = "label1";
            this.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtCode
            // 
            this.txtCode.BackColor = System.Drawing.Color.White;
            this.txtCode.Font = new System.Drawing.Font("游ゴシック", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtCode.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtCode.Location = new System.Drawing.Point(829, 61);
            this.txtCode.MaxLength = 8;
            this.txtCode.Name = "txtCode";
            this.txtCode.Size = new System.Drawing.Size(130, 50);
            this.txtCode.TabIndex = 2;
            this.txtCode.Text = "12345678";
            this.txtCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtCode.WordWrap = false;
            this.txtCode.TextChanged += new System.EventHandler(this.txtCode_TextChanged);
            this.txtCode.Enter += new System.EventHandler(this.txtYear_Enter);
            this.txtCode.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            this.txtCode.Leave += new System.EventHandler(this.txtYear_Leave);
            // 
            // lblPage
            // 
            this.lblPage.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblPage.Location = new System.Drawing.Point(1141, 877);
            this.lblPage.Name = "lblPage";
            this.lblPage.Size = new System.Drawing.Size(85, 18);
            this.lblPage.TabIndex = 278;
            this.lblPage.Text = "1000/1000";
            this.lblPage.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // chkKakunin
            // 
            this.chkKakunin.AutoSize = true;
            this.chkKakunin.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkKakunin.ForeColor = System.Drawing.SystemColors.ControlText;
            this.chkKakunin.Location = new System.Drawing.Point(1075, 877);
            this.chkKakunin.Name = "chkKakunin";
            this.chkKakunin.Size = new System.Drawing.Size(68, 18);
            this.chkKakunin.TabIndex = 5;
            this.chkKakunin.Text = "確認済";
            this.chkKakunin.UseVisualStyleBackColor = true;
            // 
            // txtMemo
            // 
            this.txtMemo.Font = new System.Drawing.Font("ＭＳ ゴシック", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMemo.ForeColor = System.Drawing.Color.DarkBlue;
            this.txtMemo.Location = new System.Drawing.Point(638, 873);
            this.txtMemo.Name = "txtMemo";
            this.txtMemo.Size = new System.Drawing.Size(422, 23);
            this.txtMemo.TabIndex = 4;
            this.txtMemo.Text = "あいうえおかきくけこ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("游ゴシック", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(827, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 281;
            this.label1.Text = "社員コード";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("游ゴシック", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.Location = new System.Drawing.Point(958, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(27, 13);
            this.label4.TabIndex = 282;
            this.label4.Text = "氏名";
            // 
            // dGV
            // 
            this.dGV.AccessibleRole = System.Windows.Forms.AccessibleRole.Grip;
            this.dGV.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(249)))), ((int)(((byte)(249)))), ((int)(((byte)(249)))));
            this.dGV.Location = new System.Drawing.Point(638, 102);
            this.dGV.Name = "dGV";
            this.dGV.Size = new System.Drawing.Size(587, 766);
            this.dGV.TabIndex = 3;
            this.dGV.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellEnter_1);
            this.dGV.CellLeave += new System.Windows.Forms.DataGridViewCellEventHandler(this.dGV_CellLeave);
            this.dGV.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dataGridView1_CellPainting);
            this.dGV.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
            this.dGV.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dataGridView1_EditingControlShowing);
            // 
            // frmCorrect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(1235, 946);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtYear);
            this.Controls.Add(this.txtMonth);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.txtMemo);
            this.Controls.Add(this.chkKakunin);
            this.Controls.Add(this.lblPage);
            this.Controls.Add(this.txtCode);
            this.Controls.Add(this.hScrollBar1);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnDataMake);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnErrCheck);
            this.Controls.Add(this.btnEnd);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.btnBefore);
            this.Controls.Add(this.btnFirst);
            this.Controls.Add(this.dGV);
            this.Controls.Add(this.btnPlus);
            this.Controls.Add(this.btnMinus);
            this.Controls.Add(this.lblNoImage);
            this.Controls.Add(this.leadImg);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.pictureBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmCorrect";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "出勤簿 兼 賃金計算書データ登録";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmCorrect_FormClosing);
            this.Load += new System.EventHandler(this.frmCorrect_Load);
            this.Shown += new System.EventHandler(this.frmCorrect_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dGV)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtMonth;
        private System.Windows.Forms.TextBox txtYear;
        private System.Windows.Forms.HScrollBar hScrollBar1;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnDataMake;
        private System.Windows.Forms.Button btnRtn;
        private System.Windows.Forms.Button btnErrCheck;
        private System.Windows.Forms.Button btnEnd;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnBefore;
        private System.Windows.Forms.Button btnFirst;
        private DataGridViewEx dGV;
        private System.Windows.Forms.Button btnPlus;
        private System.Windows.Forms.Button btnMinus;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label lblNoImage;
        private System.Windows.Forms.PictureBox pictureBox1;
        private Leadtools.WinForms.RasterImageViewer leadImg;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblErrMsg;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.TextBox txtCode;
        private System.Windows.Forms.Label lblPage;
        private System.Windows.Forms.CheckBox chkKakunin;
        private System.Windows.Forms.TextBox txtMemo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button1;
    }
}