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
using MyLibrary;


namespace CONYX_OCR.OCR
{
    public partial class frmEditLogRep : Form
    {
        public frmEditLogRep()
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
        
        string sTittle = "編集ログ一覧";

        string b_img = string.Empty;        // 印刷用画像名 2019/12/13
        bool ValChangeStatus = false;       // 2019/12/13

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        CONYXDataSet dts = new CONYXDataSet();
        CONYXDataSetTableAdapters.編集ログ表示用TableAdapter adp = new CONYXDataSetTableAdapters.編集ログ表示用TableAdapter();

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
        private string ColCheck = "c13";    // 2019/12/13
        private string ColYear = "c11";
        private string ColMonth = "c12";
        private string ColCode = "c3";
        private string ColName = "c4";
        private string ColDate = "c0";
        private string ColField = "c1";
        private string ColBefore = "c2";
        private string ColAfter = "c5";
        private string ColDateTime = "c6";
        private string ColImg = "c14";      // 2019/12/13

        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        private void GridviewSet(DataGridView tempDGV)
        {
            try
            {
                ValChangeStatus = false;    // 2019/12/13
                    
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

                // 各列設定
                //DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                //chk.Name = ColCheck;
                //tempDGV.Columns.Add(chk);
                //tempDGV.Columns[ColCheck].HeaderText = "印刷"; // 2019/12/13

                tempDGV.Columns.Add(ColYear, "年");
                tempDGV.Columns.Add(ColMonth, "月");
                tempDGV.Columns.Add(ColCode, "社員番号");
                tempDGV.Columns.Add(ColName, "氏名");
                tempDGV.Columns.Add(ColDate, "出勤簿日付");
                tempDGV.Columns.Add(ColField, "項目");
                tempDGV.Columns.Add(ColBefore, "編集前");
                tempDGV.Columns.Add(ColAfter, "編集後");
                tempDGV.Columns.Add(ColDateTime, "処理日");
                tempDGV.Columns.Add(ColImg, ""); // 2019/12/13

                // 各列幅指定
                //tempDGV.Columns[ColCheck].Width = 50; // 2019/12/13
                tempDGV.Columns[ColYear].Width = 70;
                tempDGV.Columns[ColMonth].Width = 50;
                tempDGV.Columns[ColCode].Width = 80;
                tempDGV.Columns[ColName].Width = 120;
                tempDGV.Columns[ColDate].Width = 110;
                tempDGV.Columns[ColField].Width = 100;
                tempDGV.Columns[ColBefore].Width = 100;
                tempDGV.Columns[ColAfter].Width = 100;
                tempDGV.Columns[ColDateTime].Width = 160;

                //tempDGV.Columns[ColID].Visible = false;
                tempDGV.Columns[ColImg].Visible = false;

                tempDGV.Columns[ColName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                //tempDGV.Columns[ColCheck].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // 2019/12/13
                tempDGV.Columns[ColYear].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColMonth].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColField].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColBefore].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColAfter].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColDateTime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 編集可否：2019/12/13
                tempDGV.ReadOnly = true;
                //foreach (DataGridViewColumn item in tempDGV.Columns)
                //{
                //    // チェックボックスのみ使用可
                //    if (item.Name == ColCheck)
                //    {
                //        tempDGV.Columns[item.Name].ReadOnly = false;
                //    }
                //    else
                //    {
                //        tempDGV.Columns[item.Name].ReadOnly = true;
                //    }
                //}

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

            //button2.Enabled = false;    // 2019/12/13
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

            // 過去データ読み込み
            adp.Fill(dts.編集ログ表示用, Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text));

            StringBuilder sb = new StringBuilder();

            // データグリッドビューの表示を初期化する
            g.Rows.Clear();

            var s = dts.編集ログ表示用.OrderBy(a => a.社員番号).ThenBy(a => a.年月日時刻);
            
            if (txtShainNum.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.社員番号 == txtShainNum.Text).OrderBy(a => a.社員番号).ThenBy(a => a.年月日時刻);
            }

            if (txtName.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.社員名.Contains(txtName.Text)).OrderBy(a => a.社員番号).ThenBy(a => a.年月日時刻);
            }

            ValChangeStatus = true;

            foreach (var t in s)
            {
                g.Rows.Add();
                //g[ColCheck, g.Rows.Count - 1].Value = false;
                g[ColYear, g.Rows.Count - 1].Value = t.年.ToString();
                g[ColMonth, g.Rows.Count - 1].Value = t.月.ToString();
                g[ColCode, g.Rows.Count - 1].Value = t.社員番号.ToString().PadLeft(8, '0');                
                g[ColName, g.Rows.Count - 1].Value = t.社員名;
                g[ColDate, g.Rows.Count - 1].Value = t.編集日付.ToShortDateString();
                g[ColField, g.Rows.Count - 1].Value = t.項目名;

                if (t.項目名 == "申請")
                {
                    // 変更前
                    if (t.変更前 == global.FLGOFF)
                    {
                        g[ColBefore, g.Rows.Count - 1].Value = "アンチェック";
                    }
                    else if (t.変更前 == global.FLGON)
                    {
                        g[ColBefore, g.Rows.Count - 1].Value = "チェック";
                    }
                    else if (t.変更前 == "True")
                    {
                        g[ColBefore, g.Rows.Count - 1].Value = "チェック";
                    }
                    else if (t.変更前 == "False")
                    {
                        g[ColBefore, g.Rows.Count - 1].Value = "アンチェック";
                    }

                    // 変更後
                    if (t.変更後 == global.FLGOFF)
                    {
                        g[ColAfter, g.Rows.Count - 1].Value = "アンチェック";
                    }
                    else if (t.変更後 == global.FLGON)
                    {
                        g[ColAfter, g.Rows.Count - 1].Value = "チェック";
                    }
                    else if (t.変更後 == "True")
                    {
                        g[ColAfter, g.Rows.Count - 1].Value = "チェック";
                    }
                    else if (t.変更後 == "False")
                    {
                        g[ColAfter, g.Rows.Count - 1].Value = "アンチェック";
                    }
                }
                else
                {
                    g[ColBefore, g.Rows.Count - 1].Value = t.変更前;
                    g[ColAfter, g.Rows.Count - 1].Value = t.変更後;
                }

                g[ColDateTime, g.Rows.Count - 1].Value = t.年月日時刻;
                g[ColImg, g.Rows.Count - 1].Value = t.画像名;     // 2019/12/13
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
            ImageView(dataGridView1.SelectedRows[0].Index);
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

        private void ImageView(int r)
        {
            string imgPath = Utility.NulltoStr(dataGridView1[ColImg, r].Value);

            if (imgPath.Length < 6)
            {
                MessageBox.Show("画像名が正しく登録されていません","画像表示不可",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            frmImgView frm = new frmImgView(imgPath);
            frm.ShowDialog();
        }


        ///-----------------------------------------------------------
        /// <summary>
        ///     出勤簿画像印刷 : 2019/12/13 </summary>
        ///-----------------------------------------------------------
        private void batch_ImagePrint()
        {
            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                // チェックされているスタッフを対象とする
                if (dataGridView1[ColCheck, r.Index].Value.ToString() == "False")
                {
                    continue;
                }

                // 画像名
                string img = Utility.NulltoStr(dataGridView1[ColImg, r.Index].Value);

                if (img.Length >= 6)
                {
                    // 画像フォルダパス
                    string dirName = Properties.Settings.Default.tifPath + img.Substring(0, 6) + @"\";

                    b_img = dirName + img;

                    if (System.IO.File.Exists(b_img))
                    {
                        printDocument1.Print();

                        // 後片付け
                        printDocument1.Dispose();
                    }
                }
            }
        }


        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Image img;

            img = Image.FromFile(b_img);
            e.Graphics.DrawImage(img, 0, 0);
            e.HasMorePages = false;

            img.Dispose();  // 後片付け 2019/01/15
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //batch_ImagePrint();

            ImageView(dataGridView1.SelectedRows[0].Index);
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.IsCurrentCellDirty)
            {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, "編集ログ一覧");
        }
    }
}
