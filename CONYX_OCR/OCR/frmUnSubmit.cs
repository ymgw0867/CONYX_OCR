using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using CONYX_OCR.common;


namespace CONYX_OCR.OCR
{
    public partial class frmUnSubmit : Form
    {
        public frmUnSubmit()
        {
            InitializeComponent();

            Utility.WindowsMaxSize(this, Width, Height);
            Utility.WindowsMinSize(this, Width, Height);

            //hAdp.Fill(dts.過去勤務票ヘッダ);
            //iAdp.Fill(dts.過去勤務票明細);

            txtYear.AutoSize = false;
            txtYear.Height = 25;
            txtMonth.AutoSize = false;
            txtMonth.Height = 25;
            txtShainNum.AutoSize = false;
            txtShainNum.Height = 25;
            txtName.AutoSize = false;
            txtName.Height = 25;
        }
        
        string sTittle = "過去出勤簿一覧";

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        CONYXDataSet dts = new CONYXDataSet();
        CONYXDataSetTableAdapters.過去勤務票ヘッダTableAdapter hAdp = new CONYXDataSetTableAdapters.過去勤務票ヘッダTableAdapter();
        CONYXDataSetTableAdapters.過去勤務票明細TableAdapter iAdp = new CONYXDataSetTableAdapters.過去勤務票明細TableAdapter();

        private void frmUnSubmit_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);
            
            // データグリッド定義
            GridviewSet(dataGridView1);

            // 画面初期化
            DispClear();

            // タイトル
            this.Text = sTittle;
        }
        
        // カラム定義
        private string ColDate = "c0";
        private string ColSz = "c1";
        private string ColSznm = "c2";
        private string ColCode = "c3";
        private string ColName = "c4";
        private string ColMemo = "c5";
        private string Colet = "c6";
        private string Colzn = "c7";
        private string Colsi = "c8";
        private string ColID = "c9";
        private string ColKinmuCode = "c10";
        private string ColYear = "c11";
        private string ColMonth = "c12";

        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        private void GridviewSet(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(243, 232, 207);
                //tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(20, 82, 112);

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", (float)9.5, FontStyle.Regular);

                //tempDGV.DefaultCellStyle.ForeColor = Color.FromArgb(20, 82, 112);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                //tempDGV.Height = 530;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.Control;

                // 各列幅指定
                tempDGV.Columns.Add(ColYear, "年");
                tempDGV.Columns.Add(ColMonth, "月");
                tempDGV.Columns.Add(ColCode, "社員番号");
                tempDGV.Columns.Add(ColName, "氏名");
                tempDGV.Columns.Add(ColMemo, "処理日");
                tempDGV.Columns.Add(ColID, "hID");

                tempDGV.Columns[ColYear].Width = 80;
                tempDGV.Columns[ColMonth].Width = 60;
                tempDGV.Columns[ColCode].Width = 100;
                tempDGV.Columns[ColName].Width = 170;
                tempDGV.Columns[ColMemo].Width = 180;
                tempDGV.Columns[ColID].Width = 160;

                //tempDGV.Columns[ColID].Visible = false;

                tempDGV.Columns[ColName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[ColYear].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColMonth].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColMemo].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 編集可否
                tempDGV.ReadOnly = true;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 画面初期化
        /// </summary>
        private void DispClear()
        {
            txtYear.Text = global.cnfYear.ToString();
            txtMonth.Text = global.cnfMonth.ToString();
            txtShainNum.Text = string.Empty;
        }

        private void btnSel_Click(object sender, EventArgs e)
        {
        }

        private bool errCheck()
        {
            try 
	        {
                if (txtYear.Text == string.Empty)
                {
                    txtYear.Focus();
                    throw new Exception("年を指定してください");
                }

                if (!Utility.NumericCheck(txtYear.Text))
                {
                    txtYear.Focus();
                    throw new Exception("年が正しくありません");
                }

                if (txtMonth.Text == string.Empty)
                {
                    txtMonth.Focus();
                    throw new Exception("月を指定してください");
                }

                if (!Utility.NumericCheck(txtMonth.Text))
                {
                    txtMonth.Focus();
                    throw new Exception("月が正しくありません");
                }

                if (int.Parse(txtMonth.Text) < 1 || int.Parse(txtMonth.Text) > 12)
                {
                    txtMonth.Focus();
                    throw new Exception("月が正しくありません");
                }
            }
	        catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
	        }
            return true;
        }

        /// <summary>
        /// データ検索
        /// </summary>
        private void DataSelect(DataGridView g)
        {
            this.Cursor = Cursors.WaitCursor;

            //// 奉行SQLServer接続文字列取得
            //string sc = sqlControl.obcConnectSting.get(_DBName);
            //sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 過去データ読み込み
            hAdp.FillByYYMM(dts.過去勤務票ヘッダ, Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text));

            StringBuilder sb = new StringBuilder();

            // データグリッドビューの表示を初期化する
            g.Rows.Clear();

            var s = dts.過去勤務票ヘッダ.OrderByDescending(a => a.年).ThenByDescending(a => a.月).ThenBy(a => a.社員番号).ThenByDescending(a => a.更新年月日);
            
            //if (txtYear.Text.Trim() != string.Empty)
            //{
            //    s = s.Where(a => a.年 == Utility.StrtoInt(txtYear.Text))
            //        .OrderByDescending(a => a.年).ThenByDescending(a => a.月).ThenBy(a => a.社員番号).ThenByDescending(a => a.更新年月日);
            //}

            //if (txtMonth.Text.Trim() != string.Empty)
            //{
            //    s = s.Where(a => a.月 == Utility.StrtoInt(txtMonth.Text))
            //       .OrderByDescending(a => a.年).ThenByDescending(a => a.月).ThenBy(a => a.社員番号).ThenByDescending(a => a.更新年月日);
            //}

            if (txtShainNum.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.社員番号 == Utility.StrtoInt(txtShainNum.Text))
                   .OrderByDescending(a => a.年).ThenByDescending(a => a.月).ThenBy(a => a.社員番号).ThenByDescending(a => a.更新年月日);
            }

            if (txtName.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.社員名.Contains(txtName.Text))
                   .OrderByDescending(a => a.年).ThenByDescending(a => a.月).ThenBy(a => a.社員番号).ThenByDescending(a => a.更新年月日);
            }

            foreach (var t in s)
            {
                g.Rows.Add();
                g[ColYear, g.Rows.Count - 1].Value = t.年.ToString();
                g[ColMonth, g.Rows.Count - 1].Value = t.月.ToString();
                g[ColCode, g.Rows.Count - 1].Value = t.社員番号.ToString().PadLeft(8, '0');                
                g[ColName, g.Rows.Count - 1].Value = t.社員名;
                g[ColMemo, g.Rows.Count - 1].Value = t.更新年月日.ToShortDateString() + " " + t.更新年月日.Hour.ToString().PadLeft(2, '0') + ":" + t.更新年月日.Minute.ToString().PadLeft(2, '0') + ":" + t.更新年月日.Second.ToString().PadLeft(2, '0');

                g[ColID, g.Rows.Count - 1].Value = t.ID.ToString();
            }
            
            dataGridView1.CurrentCell = null;

            // 終了
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("該当するデータはありませんでした", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                this.Text = sTittle + " " + dataGridView1.RowCount.ToString("#,##0") + "件"; 
            }

            this.Cursor = Cursors.Default;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmUnSubmit_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string rID = string.Empty;

            rID = dataGridView1[ColID, dataGridView1.SelectedRows[0].Index].Value.ToString();

            if (rID != string.Empty)
            {
                this.Hide();
                OCR.frmPastCorrect frm = new frmPastCorrect(rID);
                frm.ShowDialog();
                this.Show();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (errCheck())
            {
                DataSelect(dataGridView1);
            }
        }

        private void txtRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
