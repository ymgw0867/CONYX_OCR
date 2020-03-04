using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CONYX_OCR.common;

namespace CONYX_OCR.OCR
{
    public partial class frmOCRPC : Form
    {
        public frmOCRPC()
        {
            InitializeComponent();
        }

        private void frmOCRPC_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最少サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // コンボボックス
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;

            // コンボボックス
            loadOutPcMst();
        }

        public string _outPC { get; set; }
        
        ///---------------------------------------------------
        /// <summary>
        ///     出力先ＰＣコンボボックスへロードする</summary>
        /// 
        ///---------------------------------------------------
        private void loadOutPcMst()
        {
            CONYXDataSet dts = new CONYXDataSet();
            CONYXDataSetTableAdapters.出力先ＰＣTableAdapter adp = new CONYXDataSetTableAdapters.出力先ＰＣTableAdapter();

            adp.Fill(dts.出力先ＰＣ);

            var s = dts.出力先ＰＣ.OrderBy(a => a.ID);

            foreach (var t in s)
            {
                comboBox1.Items.Add(t.登録名);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == string.Empty)
            {
                MessageBox.Show("出力先ＰＣを選択してください", "出力先ＰＣ未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 出力先ＰＣ登録名取得
            _outPC = comboBox1.Text;

            // 入力画像があるか？
            var tifCnt = System.IO.Directory.GetFiles(Properties.Settings.Default.scanPath, "*.tif").Count();

            if (tifCnt == 0)
            {
                MessageBox.Show("出勤簿画像がありません", "画像なし", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                _outPC = string.Empty;

                // フォームを閉じる
                this.Close();
            }
            
            // フォームを閉じる
            this.Close();
        }

        private void frmOCRPC_FormClosing(object sender, FormClosingEventArgs e)
        {
            //this.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            _outPC = string.Empty;

            // フォームを閉じる
            this.Close();
        }
    }
}
