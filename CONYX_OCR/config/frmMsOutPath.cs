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
using CONYX_OCR.config;

namespace CONYX_OCR.config
{
    public partial class frmMsOutPath : Form
    {
        public frmMsOutPath()
        {
            InitializeComponent();
        }

        CONYXDataSet dts = new CONYXDataSet();
        CONYXDataSetTableAdapters.出力先ＰＣTableAdapter adp = new CONYXDataSetTableAdapters.出力先ＰＣTableAdapter();

        private void frmGekkyuKinmu_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);  // フォーム最大サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);  // フォーム最小サイズ
            GridViewSetting(dg1);                                   // グリッドビュー設定
            GridViewShow(dg1);                                      // グリッドビューへデータ表示
            DispClr();                                              // 画面初期化
        }

        // データ領域情報
        string _dbName = string.Empty;
        string _comName = string.Empty;

        const int ADDMODE = 0;
        const int UPMODE = 1;

        // 登録モード
        int _fMode = ADDMODE;

        // 選択データＩＤ
        string _SelectID = "0";

        // 登録内容
        string _Pt = string.Empty;
        string _BumonCode = string.Empty;   // 部門コード
        string _BumonName = string.Empty;   // 部門名

        // グリッドビューカラム名
        private string cCode = "c1";
        private string cComputerName = "c2";
        private string cOutPath = "c3";
        private string cName = "c4";
        private string cBikou = "c5";
        private string cDate = "c6";

        /// <summary>
        /// グリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.BurlyWood;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", (float)9.5, FontStyle.Regular);

                // 行の高さ
                //tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 149;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.Control;

                //各列幅指定
                tempDGV.Columns.Add(cCode, "コード");
                tempDGV.Columns.Add(cComputerName, "コンピュータ名");
                tempDGV.Columns.Add(cName, "登録名");
                tempDGV.Columns.Add(cBikou, "備考");
                tempDGV.Columns.Add(cDate, "更新日");

                tempDGV.Columns[cCode].Width = 74;
                tempDGV.Columns[cComputerName].Width = 140;
                tempDGV.Columns[cName].Width = 140;
                //tempDGV.Columns[cBikou].Width = 120;
                tempDGV.Columns[cBikou].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[cDate].Width = 110;

                tempDGV.Columns[cCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更不可
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                tempDGV.StandardTab = false;

                //// ソート禁止
                //foreach (DataGridViewColumn c in tempDGV.Columns)
                //{
                //    c.SortMode = DataGridViewColumnSortMode.NotSortable;
                //}
                //tempDGV.Columns[cDay].SortMode = DataGridViewColumnSortMode.NotSortable;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        /// <summary>
        ///     データをグリッドビューへ表示します </summary>
        /// <param name="tempGrid">
        ///     データグリッドビューオブジェクト</param>
        private void GridViewShow(DataGridView tempGrid)
        {
            // データ読み込み
            adp.Fill(dts.出力先ＰＣ);

            // データ表示
            int iX = 0;
            tempGrid.RowCount = 0;

            var s = dts.出力先ＰＣ.OrderBy(a => a.ID);

            foreach (var t in s)
            {
                tempGrid.Rows.Add();

                tempGrid[cCode, iX].Value = t.ID.ToString();
                tempGrid[cComputerName, iX].Value = t.コンピューター名;
                tempGrid[cName, iX].Value = t.登録名;
                tempGrid[cBikou, iX].Value = t.備考;
                tempGrid[cDate, iX].Value = DateTime.Parse(t.更新年月日.ToString()).ToShortDateString();

                iX++;
            }

            tempGrid.CurrentCell = null;
           
        }

        private void DispClr()
        {
            txtComName.Text = System.Net.Dns.GetHostName();
            txtName.Text = string.Empty;
            txtBikou.Text = string.Empty;

            button2.Enabled = true;
            button3.Enabled = false;
            button5.Enabled = false;

            dg1.CurrentCell = null;

            _fMode = ADDMODE;   // 処理モード
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmGekkyuKinmu_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void dg1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // 選択データ表示
            GetGridData(e.RowIndex);

            // ボタン
            button2.Enabled = true;
            button3.Enabled = true;
            button5.Enabled = true;

            // 処理モード
            _fMode = UPMODE;
        }

        /// <summary>
        /// 選択データ表示
        /// </summary>
        /// <param name="r">グリッドビュー行Index</param>
        private void GetGridData(int r)
        {
            _SelectID = dg1[cCode, r].Value.ToString();

            txtComName.Text = dg1[cComputerName, r].Value.ToString();
            txtName.Text = dg1[cName, r].Value.ToString();
            txtBikou.Text = dg1[cBikou, r].Value.ToString();
            txtName.Focus();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (txtName.Text.Trim() == string.Empty)
            {
                MessageBox.Show("登録名を入力して下さい","登録名未入力",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                txtName.Focus();
                return;
            }

            if (nameSelect())
            {
                MessageBox.Show("この登録名は既に登録済みです", "登録済み", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtName.Focus();
                return;
            }

            if (_fMode == ADDMODE)
            {
                if (comNameSelect())
                {
                    MessageBox.Show("このコンピュータは既に登録済みです", "登録済み", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }

            if (_fMode == UPMODE)
            {
                if (txtComName.Text != System.Net.Dns.GetHostName())
                {
                    if (MessageBox.Show("他のコンピュータの情報を編集しようとしています。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.No)
                        return;
                }
            }

            if (MessageBox.Show("このコンピュータをＯＣＲ出力先へ登録します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.No)
                return;

            // データ更新
            if (_fMode == ADDMODE)
            {
                dataAdd();
            }

            if (_fMode == UPMODE)
            {
                dataUpdate();
            }

            // データグリッド再表示
            GridViewShow(dg1);

            // 画面初期化
            DispClr();
        }

        private void dataAdd()
        {
            try
            {
                dts.出力先ＰＣ.Add出力先ＰＣRow(txtComName.Text, txtName.Text, txtBikou.Text, 0, DateTime.Now);

                adp.Update(dts.出力先ＰＣ);

                // ＰＣ登録名　2014/05/20
                global.pcName = txtName.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            finally
            {
                //if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// データ更新
        /// </summary>
        private void dataUpdate()
        {
            try
            {
                CONYXDataSet.出力先ＰＣRow r = dts.出力先ＰＣ.Single(a => a.ID == int.Parse(_SelectID));

                r.コンピューター名 = txtComName.Text;
                r.登録名 = txtName.Text;
                r.備考 = txtBikou.Text;
                r.更新年月日 = DateTime.Now;

                adp.Update(dts.出力先ＰＣ);

                // 自分のコンピュータ名のときＰＣ登録名　2014/05/20
                if (System.Net.Dns.GetHostName() == txtComName.Text)
                {
                    global.pcName = txtName.Text;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            finally
            {
                //if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DispClr();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("削除します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.No)
                return;

            // データ削除
            dataDel();

            // 画面再表示
            GridViewShow(dg1);

            // 画面初期化
            DispClr();
        }

        private void dataDel()
        {
            try
            {
                CONYXDataSet.出力先ＰＣRow r = dts.出力先ＰＣ.Single(a => a.ID == int.Parse(_SelectID));
                r.Delete();
                adp.Update(dts.出力先ＰＣ);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            finally
            {
                //if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        ///     コンピュータ名検索</summary>
        /// <returns>
        ///     true:登録済み、false:未登録</returns>
        private bool comNameSelect()
        {
            bool result = true;

            try
            {
                int s = dts.出力先ＰＣ.Where(a => a.コンピューター名 == txtComName.Text).Count();

                if (s > 0)
                {
                    result = true;
                }
                else
                {
                    result = false;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
            
            return result;
        }

        /// <summary>
        ///     登録名検索</summary>
        /// <returns>
        ///     true:登録済み、false:未登録</returns>
        private bool nameSelect()
        {
            bool result = true;

            try
            {
                int s = dts.出力先ＰＣ.Where(a => a.登録名 == txtName.Text && a.コンピューター名 != txtComName.Text).Count();

                if (s > 0)
                {
                    result = true;
                }
                else
                {
                    result = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }

            return result;
        }

        private void dg1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
