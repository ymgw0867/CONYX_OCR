using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using CONYX_OCR.common;

namespace CONYX_OCR.OCR
{
    public partial class frmPastCorrect : Form
    {
        /// ------------------------------------------------------------
        /// <summary>
        ///     コンストラクタ </summary>
        /// <param name="dbName">
        ///     会社領域データベース名</param>
        /// <param name="comName">
        ///     会社名</param>
        /// <param name="sID">
        ///     処理モード</param>
        /// ------------------------------------------------------------
        public frmPastCorrect(string sID)
        {
            InitializeComponent();
            dID = sID;              // 処理モード

            // テーブルアダプターマネージャーに過去勤務票ヘッダ、明細テーブルアダプターを割り付ける
            adpMn.過去勤務票ヘッダTableAdapter = hAdp;
            adpMn.過去勤務票明細TableAdapter = iAdp;

            txtYear.AutoSize = false;
            txtYear.Height = 48;
            txtMonth.AutoSize = false;
            txtMonth.Height = 48;
            txtCode.AutoSize = false;
            txtCode.Height = 36;
            txtMemo.AutoSize = false;
            txtMemo.Height = 24;
        }

        // データアダプターオブジェクト
        CONYXDataSetTableAdapters.TableAdapterManager adpMn = new CONYXDataSetTableAdapters.TableAdapterManager();
        CONYXDataSetTableAdapters.過去勤務票ヘッダTableAdapter hAdp = new CONYXDataSetTableAdapters.過去勤務票ヘッダTableAdapter();
        CONYXDataSetTableAdapters.過去勤務票明細TableAdapter iAdp = new CONYXDataSetTableAdapters.過去勤務票明細TableAdapter();
        CONYXDataSetTableAdapters.出勤簿編集ログTableAdapter gAdp = new CONYXDataSetTableAdapters.出勤簿編集ログTableAdapter();


        // データセットオブジェクト
        CONYXDataSet dts = new CONYXDataSet();
        CONYX_CLIDataSet dtsCl = new CONYX_CLIDataSet();

        // セル値
        private string cellName = string.Empty;         // セル名
        private string cellBeforeValue = string.Empty;  // 編集前
        private string cellAfterValue = string.Empty;   // 編集後

        #region 編集ログ・項目名
        private const string LOG_YEAR = "年";
        private const string LOG_MONTH = "月";
        private const string LOG_NUMBER = "社員番号";
        private const string CELL_TAIKEICD = "勤務体系";
        private const string CELL_CHECK = "申請";
        private const string CELL_KAISHI = "出勤・時";
        private const string CELL_KAISHI_M = "出勤・分";
        private const string CELL_TAIKIN = "退勤・時";
        private const string CELL_TAIKIN_M = "退勤・分";
        private const string CELL_REST = "休憩・時";
        private const string CELL_REST_M = "休憩・分";
        private const string CELL_WORK = "実労・時";
        private const string CELL_WORK_M = "実労・分";
        private const string CELL_JIYU1 = "事由１";
        private const string CELL_JIYU2 = "事由２";
        #endregion 編集ログ・項目名

        // カレント社員情報
        //CONYXDataSet.社員所属Row cSR = null;

        // カレントデータRowsインデックス
        string[] cIdx = null;

        /// <summary>
        ///     カレントデータRowsインデックス</summary>
        int cI = 0;

        // 社員マスターより取得した所属コード
        string mSzCode = string.Empty;

        #region 終了ステータス定数
        const string END_BUTTON = "btn";
        const string END_MAKEDATA = "data";
        const string END_CONTOROL = "close";
        const string END_NODATA = "non Data";
        #endregion

        string dID = string.Empty;              // 表示する過去データのID
        string sDBNM = string.Empty;            // データベース名

        string _dbName = string.Empty;          // 会社領域データベース識別番号
        string _comNo = string.Empty;           // 会社番号
        string _comName = string.Empty;         // 会社名

        int dtIndi = global.flgOff;             // 就業奉行時間表記設定取得（0：24時間表記 1：48時間表記）

        // dataGridView1_CellEnterステータス
        bool gridViewCellEnterStatus = true;

        // 編集ログ書き込み状態
        bool editLogStatus = false;

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // 自分のコンピュータの登録名を取得
            // 登録されていないとき終了します
            string pcName = Utility.getPcDir();
            if (pcName == string.Empty)
            {
                MessageBox.Show("このコンピュータがＯＣＲ出力先として登録されていません。", "出力先未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.Close();
            }

            // スキャンＰＣのコンピュータ別フォルダ内のＯＣＲデータ存在チェック
            if (Directory.Exists(Properties.Settings.Default.pcPath + pcName + @"\"))
            {
                string[] ocrfiles = Directory.GetFiles(Properties.Settings.Default.pcPath + pcName + @"\", "*.csv");

                // スキャンＰＣのＯＣＲ画像、ＣＳＶデータをローカルのDATAフォルダへ移動します
                if (ocrfiles.Length > 0)
                {
                    foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.pcPath + pcName + @"\", "*"))
                    {
                        // パスを含まないファイル名を取得
                        string reFile = Path.GetFileName(files);

                        // ファイル移動
                        if (reFile != "Thumbs.db")
                        {
                            File.Move(files, Properties.Settings.Default.dataPath + @"\" + reFile);
                        }
                    }
                }
            }
            
            // データセットへデータを読み込みます
            getDataSet(dID);

            // キャプション
            //this.Text = "出勤簿データ表示";

            // グリッドビュー定義
            GridviewSet gs = new GridviewSet();
            gs.Setting_Shain(dGV);

            // レコードを表示
            showOcrData(dID);

            // tagを初期化
            this.Tag = string.Empty;
        }

        #region データグリッドビューカラム定義
        private static string cCheck = "col1";          // 取消
        private static string cDay = "col17";           // 日
        private static string cKinmuTaikei = "col2";    // 勤務体系コード
        private static string cFlg = "col3";        // 申請フラグ
        private static string cZH = "col5";         // 残業時
        private static string cZE = "col6";         // :
        private static string cZM = "col7";         // 残業分
        private static string cSIH = "col8";        // 深夜時
        private static string cSIE = "col9";        // :
        private static string cSIM = "col10";       // 深夜分
        private static string cSH = "col11";        // 開始時
        private static string cSE = "col12";        // :
        private static string cSM = "col13";        // 開始分
        private static string cEH = "col14";        // 終了時
        private static string cEE = "col15";        // :
        private static string cEM = "col16";        // 終了分
        private static string cRH = "col18";        // 休憩時
        private static string cRE = "col19";        // :
        private static string cRM = "col20";        // 休憩分
        private static string cWH = "col21";        // 実労時
        private static string cWE = "col22";        // :
        private static string cWM = "col23";        // 実労分
        private static string cJ1 = "col24";        // 事由１
        private static string cJ2 = "col25";        // 事由２
        private static string cWeek = "col26";      // 曜日
        private static string cID = "colID";        // ID

        #endregion
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビュークラス </summary>
        ///----------------------------------------------------------------------------
        private class GridviewSet
        {
            ///----------------------------------------------------------------------------
            /// <summary>
            ///     社員用データグリッドビューの定義を行います</summary> 
            /// <param name="gv">
            ///     データグリッドビューオブジェクト</param>
            ///----------------------------------------------------------------------------
            public void Setting_Shain(DataGridView gv)
            {
                try
                {
                    // データグリッドビューの基本設定
                    setGridView_Properties(gv);

                    // カラムコレクションを空にします
                    gv.Columns.Clear();

                    // 行数をクリア            
                    gv.Rows.Clear();

                    //各列幅指定
                    gv.Columns.Add(cDay, "日");
                    gv.Columns.Add(cWeek, "曜");
                    gv.Columns.Add(cKinmuTaikei, "勤務体系");

                    DataGridViewCheckBoxColumn column = new DataGridViewCheckBoxColumn();
                    gv.Columns.Add(column);
                    gv.Columns[3].Name = cFlg;
                    gv.Columns[3].HeaderText = "申請";

                    gv.Columns.Add(cSH, "出");
                    gv.Columns.Add(cSE, "");
                    gv.Columns.Add(cSM, "勤");
                    gv.Columns.Add(cEH, "退");
                    gv.Columns.Add(cEE, "");
                    gv.Columns.Add(cEM, "勤");
                    gv.Columns.Add(cRH, "休");
                    gv.Columns.Add(cRE, "");
                    gv.Columns.Add(cRM, "憩");
                    gv.Columns.Add(cWH, "実");
                    gv.Columns.Add(cWE, "");
                    gv.Columns.Add(cWM, "労");
                    gv.Columns.Add(cJ1, "事由１");
                    gv.Columns.Add(cJ2, "事由２");

                    gv.Columns.Add(cID, "");        // 明細ID
                    gv.Columns[cID].Visible = false;

                    foreach (DataGridViewColumn c in gv.Columns)
                    {
                        // 幅
                        if (c.Name == cDay || c.Name == cWeek)
                        {
                            c.Width = 30;
                        }
                        else if (c.Name == cKinmuTaikei)
                        {
                            c.Width = 70;
                        }
                        else if (c.Name == cFlg)
                        {
                            c.Width = 50;
                        }
                        else if (c.Name == cJ1 || c.Name == cJ2)
                        {
                            c.Width = 62;
                        }
                        else if (c.Name == cSE || c.Name == cEE || c.Name == cRE || c.Name == cWE)
                        {
                            c.Width = 10;
                        }
                        else
                        {
                            c.Width = 30;
                        }
                                                
                        // 表示位置
                        if (c.Index < 3 || c.Name == cSE || c.Name == cEE || c.Name == cRE || c.Name == cWE || 
                            c.Name == cJ1 || c.Name == cJ2)
                        {
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        }

                        if (c.Name == cSH || c.Name == cEH || c.Name == cRH || c.Name == cWH)
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        if (c.Name == cSM || c.Name == cEM || c.Name == cRM || c.Name == cWM)
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

                        // 編集可否
                        if (c.Name == cDay || c.Name == cWeek || c.Name == cSE || c.Name == cEE || c.Name == cRE || c.Name == cWE)
                        {
                            c.ReadOnly = true;
                        }
                        else
                        {
                            // 全ての項目を編集不可
                            c.ReadOnly = true;
                        }

                        // 区切り文字
                        if (c.Name == cSE || c.Name == cEE || c.Name == cRE || c.Name == cWE)
                        {
                            c.DefaultCellStyle.Font = new Font("游ゴシック", 8, FontStyle.Regular);
                        }

                        // 入力可能桁数
                        if (c.Name != cFlg)
                        {
                            DataGridViewTextBoxColumn col = (DataGridViewTextBoxColumn)c;

                            if (c.Name == cRH)
                            {
                                col.MaxInputLength = 1;
                            }
                            else if (c.Name == cKinmuTaikei)
                            {
                                col.MaxInputLength = 4;
                            }
                            else
                            {
                                col.MaxInputLength = 2;
                            }
                        }

                        if (c.Name == cDay)
                        {
                            
                        } 


                        // ソート禁止
                        c.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            ///----------------------------------------------------------------------------
            /// <summary>
            ///     データグリッドビュー基本設定</summary>
            /// <param name="gv">
            ///     データグリッドビューオブジェクト</param>
            ///----------------------------------------------------------------------------
            private void setGridView_Properties(DataGridView gv)
            {
                // 列スタイルを変更する
                gv.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                gv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                

                // 列ヘッダーフォント指定
                gv.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);
                gv.ColumnHeadersDefaultCellStyle.ForeColor = Color.DarkBlue;

                // データフォント指定
                gv.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", (Single)11, FontStyle.Regular);

                // 行の高さ
                gv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                gv.ColumnHeadersHeight = 20;
                gv.RowTemplate.Height = 24;

                // 全体の高さ
                //gv.Height = 362;
                //gv.Height = 735;

                // 全体の幅
                gv.Width = 587;

                // 奇数行の色
                //gv.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //テキストカラーの設定
                gv.RowsDefaultCellStyle.ForeColor = global.defaultColor;
                gv.DefaultCellStyle.SelectionBackColor = Color.Empty;
                gv.DefaultCellStyle.SelectionForeColor = global.defaultColor;

                // 行ヘッダを表示しない
                gv.RowHeadersVisible = false;

                // 選択モード
                gv.SelectionMode = DataGridViewSelectionMode.CellSelect;
                gv.MultiSelect = false;

                // データグリッドビュー編集可
                gv.ReadOnly = false;

                // 追加行表示しない
                gv.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                gv.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                gv.AllowUserToOrderColumns = false;

                // 列サイズ変更不可
                gv.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                gv.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //gv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                gv.StandardTab = false;

                // 編集モード
                gv.EditMode = DataGridViewEditMode.EditOnEnter;
            }
        }


        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVデータをMDBへインサートする</summary>
        ///----------------------------------------------------------------------------
        private void GetCsvDataToMDB()
        {
            // CSVファイル数をカウント
            string[] inCsv = System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath, "*.csv");

            // CSVファイルがなければ終了
            if (inCsv.Length == 0) return;

            // オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // OCRのCSVデータをMDBへ取り込む
            OCRData ocr = new OCRData(_dbName);
            ocr.CsvToMdb(Properties.Settings.Default.dataPath, frmP, _dbName);

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //if (e.Control is DataGridViewTextBoxEditingControl)
            //{
            //    // 数字のみ入力可能とする
            //    if (dGV.CurrentCell.ColumnIndex > 3)
            //    {
            //        //イベントハンドラが複数回追加されてしまうので最初に削除する
            //        e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
            //        e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);

            //        //イベントハンドラを追加する
            //        e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            //    }
            //}
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!global.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            string colName = dGV.Columns[e.ColumnIndex].Name;
            
            if (colName == cWeek)
            {
                if (Utility.NulltoStr(dGV[cWeek, e.RowIndex].Value) == "土")
                {
                    dGV[cDay, e.RowIndex].Style.BackColor = Color.AliceBlue;
                    dGV[cWeek, e.RowIndex].Style.BackColor = Color.AliceBlue;
                }
                else if (Utility.NulltoStr(dGV[cWeek, e.RowIndex].Value) == "日")
                {
                    dGV[cDay, e.RowIndex].Style.BackColor = Color.MistyRose;
                    dGV[cWeek, e.RowIndex].Style.BackColor = Color.MistyRose;
                }
                else if (Utility.NulltoStr(dGV[cWeek, e.RowIndex].Value) == string.Empty)

                    dGV.Rows[e.RowIndex].DefaultCellStyle.BackColor = SystemColors.Control;
                }
                else
                {
                {
                    dGV[cDay, e.RowIndex].Style.BackColor = SystemColors.Control;
                    dGV[cWeek, e.RowIndex].Style.BackColor = SystemColors.Control;
                    //dGV.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Empty;
                    dGV.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255, 254, 255);
                }
            }
        }
        
        private void frmCorrect_Shown(object sender, EventArgs e)
        {
            if (dID != string.Empty) btnRtn.Focus();
        }

        private void dataGridView3_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //if (e.Control is DataGridViewTextBoxEditingControl)
            //{
            //    //イベントハンドラが複数回追加されてしまうので最初に削除する
            //    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
            //    //イベントハンドラを追加する
            //    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            //}
        }

        private void dataGridView4_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //if (e.Control is DataGridViewTextBoxEditingControl)
            //{
            //    //イベントハンドラが複数回追加されてしまうので最初に削除する
            //    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
            //    //イベントハンドラを追加する
            //    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            //}
        }

        private void btnNext_Click(object sender, EventArgs e)
        {

        }
        
        ///------------------------------------------------------
        /// <summary>
        ///     暦日付を求める </summary>
        /// <param name="sYY">
        ///     対象年</param>
        /// <param name="sMM">
        ///     対象月</param>
        /// <param name="sDD">
        ///     出勤簿の日</param>
        /// <returns>
        ///     歴日付</returns>
        ///------------------------------------------------------
        private DateTime getCalenderDate(int sYY, int sMM, int sDD)
        {
            DateTime rtnDate = global.NODATE;
            DateTime dt = DateTime.Today;

            if (sDD > 15)
            {
                if (DateTime.TryParse(sYY + "/" + sMM + "/" + sDD, out dt))
                {
                    rtnDate = dt;
                }
            }
            else if ((sMM + 1) > 12)
            {
                if (DateTime.TryParse((sYY + 1) + "/01/" + sDD, out dt))
                {
                    rtnDate = dt;
                }
            }
            else if (DateTime.TryParse(sYY + "/" + (sMM + 1) + "/" + sDD, out dt))
            {
                rtnDate = dt;
            }

            return rtnDate;
        }


        /// ----------------------------------------------------------------------------------------------------
        /// <summary>
        ///     空白以外のとき、指定された文字数になるまで左側に０を埋めこみ、右寄せした文字列を返す
        /// </summary>
        /// <param name="tm">
        ///     文字列</param>
        /// <param name="len">
        ///     文字列の長さ</param>
        /// <returns>
        ///     文字列</returns>
        /// ----------------------------------------------------------------------------------------------------
        private string timeVal(object tm, int len)
        {
            string t = Utility.NulltoStr(tm);
            if (t != string.Empty) return t.PadLeft(len, '0');
            else return t;
        }

        /// ----------------------------------------------------------------------------------------------------
        /// <summary>
        ///     空白以外のとき、先頭文字が０のとき先頭文字を削除した文字列を返す　
        ///     先頭文字が０以外のときはそのまま返す
        /// </summary>
        /// <param name="tm">
        ///     文字列</param>
        /// <returns>
        ///     文字列</returns>
        /// ----------------------------------------------------------------------------------------------------
        private string timeValH(object tm)
        {
            string t = Utility.NulltoStr(tm);

            if (t != string.Empty)
            {
                t = t.PadLeft(2, '0');
                if (t.Substring(0, 1) == "0")
                {
                    t = t.Substring(1, 1);
                }
            }

            return t;
        }

        /// ------------------------------------------------------------------------------------
        /// <summary>
        ///     Bool値を数値に変換する </summary>
        /// <param name="b">
        ///     True or False</param>
        /// <returns>
        ///     true:1, false:0</returns>
        /// ------------------------------------------------------------------------------------
        private int booltoFlg(string b)
        {
            if (b == "True") return global.flgOn;
            else return global.flgOff;
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {

        }

        private void btnBefore_Click(object sender, EventArgs e)
        {

        }

        private void btnFirst_Click(object sender, EventArgs e)
        {

        }
        
        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {

        }
        
        private void btnRtn_Click(object sender, EventArgs e)
        {
            // 非ログ書き込み状態とする
            editLogStatus = false;

            // フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            //「受入データ作成終了」「勤務票データなし」以外での終了のとき
            if (this.Tag.ToString() != END_MAKEDATA && this.Tag.ToString() != END_NODATA)
            {
                if (MessageBox.Show("終了します。よろしいですか", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }
            }
            
            // 解放する
            this.Dispose();
        }
        
        /// ---------------------------------------------------------------------
        /// <summary>
        ///     MDBファイルを最適化する </summary>
        /// ---------------------------------------------------------------------
        private void mdbCompact()
        {
            try
            {
                JRO.JetEngine jro = new JRO.JetEngine();
                string OldDb = Properties.Settings.Default.mdbOlePathCL;
                string NewDb = Properties.Settings.Default.mdbPathTempCL;

                jro.CompactDatabase(OldDb, NewDb);

                //今までのバックアップファイルを削除する
                System.IO.File.Delete(Properties.Settings.Default.mdbCLPath + global.MDBBACK);

                //今までのファイルをバックアップとする
                System.IO.File.Move(Properties.Settings.Default.mdbCLPath + global.MDBFILE, Properties.Settings.Default.mdbCLPath + global.MDBBACK);

                //一時ファイルをMDBファイルとする
                System.IO.File.Move(Properties.Settings.Default.mdbCLPath + global.MDBTEMP, Properties.Settings.Default.mdbCLPath + global.MDBFILE);
            }
            catch (Exception e)
            {
                MessageBox.Show("MDB最適化中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }
        
        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < global.ZOOM_MAX)
            {
                leadImg.ScaleFactor += global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > global.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= global.ZOOM_STEP;
            }

            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }
        
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //if (e.RowIndex < 0) return;

            string colName = dGV.Columns[e.ColumnIndex].Name;

            if (colName == cSH || colName == cSE || colName == cEH || colName == cEE ||
                colName == cRH || colName == cRE || colName == cWH || colName == cWE)
            {
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            string colName = dGV.Columns[dGV.CurrentCell.ColumnIndex].Name;
            //if (colName == cKyuka || colName == cCheck)
            //{
            //    if (dGV.IsCurrentCellDirty)
            //    {
            //        dGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
            //        dGV.RefreshEdit();
            //    }
            //}
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView1_CellEnter_1(object sender, DataGridViewCellEventArgs e)
        {
            // ログ書き込み状態のとき、値を保持する
            if (editLogStatus)
            {
                // 勤務体系
                if (e.ColumnIndex == 2)
                {
                    cellName = CELL_TAIKEICD;
                }

                // 申請チェック
                if (e.ColumnIndex == 3)
                {
                    cellName = CELL_CHECK;
                }

                // 始業・時
                if (e.ColumnIndex == 4)
                {
                    cellName = CELL_KAISHI;
                }

                // 始業・分
                if (e.ColumnIndex == 6)
                {
                    cellName = CELL_KAISHI_M;
                }

                // 終業・時
                if (e.ColumnIndex == 7)
                {
                    cellName = CELL_TAIKIN;
                }

                // 終業・分
                if (e.ColumnIndex == 9)
                {
                    cellName = CELL_TAIKIN_M;
                }
                
                // 休憩・時
                if (e.ColumnIndex == 10)
                {
                    cellName = CELL_REST;
                }

                // 休憩・分
                if (e.ColumnIndex == 12)
                {
                    cellName = CELL_REST_M;
                }

                // 実労・時
                if (e.ColumnIndex == 13)
                {
                    cellName = CELL_WORK;
                }

                // 実労・分
                if (e.ColumnIndex == 15)
                {
                    cellName = CELL_WORK_M;
                }

                // 事由１
                if (e.ColumnIndex == 16)
                {
                    cellName = CELL_JIYU1;
                }

                // 事由２
                if (e.ColumnIndex == 17)
                {
                    cellName = CELL_JIYU2;
                }

                // セルの値を保持
                //if (dGV[e.ColumnIndex, e.RowIndex].Value == null)
                //{
                //    cellBeforeValue = string.Empty;
                //}
                //else
                //{
                //    cellBeforeValue = dGV[e.ColumnIndex, e.RowIndex].Value.ToString();
                //}

                cellBeforeValue = Utility.NulltoStr(dGV[e.ColumnIndex, e.RowIndex].Value);
            }

            // エラー表示時には処理を行わない
            if (!gridViewCellEnterStatus)
            {
                return;
            }
 
            string ColH = string.Empty;
            string ColM = dGV.Columns[dGV.CurrentCell.ColumnIndex].Name;

            // 開始時間または終了時間を判断
            if (ColM == cSM)            // 開始時刻
            {
                ColH = cSH;
            }
            else if (ColM == cEM)       // 終了時刻
            {
                ColH = cEH;
            }
            else if (ColM == cRM)       // 休憩
            {
                ColH = cRH;
            }
            else if (ColM == cWM)      // 実労
            {
                ColH = cWH;
            }
            else
            {
                return;
            }

            // 時が入力済みで分が未入力のとき分に"00"を表示します
            if (dGV[ColH, dGV.CurrentRow.Index].Value != null)
            {
                if (dGV[ColH, dGV.CurrentRow.Index].Value.ToString().Trim() != string.Empty)
                {
                    if (dGV[ColM, dGV.CurrentRow.Index].Value == null)
                    {
                        dGV[ColM, dGV.CurrentRow.Index].Value = "00";
                    }
                    else if (dGV[ColM, dGV.CurrentRow.Index].Value.ToString().Trim() == string.Empty)
                    {
                        dGV[ColM, dGV.CurrentRow.Index].Value = "00";
                    }
                }
            }
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     伝票画像表示 </summary>
        /// <param name="iX">
        ///     現在の伝票</param>
        /// <param name="tempImgName">
        ///     画像名</param>
        /// ------------------------------------------------------------------------------
        public void ShowImage(string tempImgName)
        {
            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                lblNoImage.Visible = true;
                global.pblImagePath = string.Empty;
                button1.Enabled = false;
                return;
            }

            //画像ファイルがあるとき表示
            if (File.Exists(tempImgName))
            {
                lblNoImage.Visible = false;
                leadImg.Visible = true;
                button1.Enabled = true;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;

                //画像ロード
                Leadtools.Codecs.RasterCodecs.Startup();
                Leadtools.Codecs.RasterCodecs cs = new Leadtools.Codecs.RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                Leadtools.RasterPaintProperties prop = new Leadtools.RasterPaintProperties();
                prop = Leadtools.RasterPaintProperties.Default;
                prop.PaintDisplayMode = Leadtools.RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(tempImgName, 0, Leadtools.Codecs.CodecsLoadByteOrder.BgrOrGray, 1, 1);

                //画像表示倍率設定
                leadImg.ScaleFactor *= global.ZOOM_RATE;

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = Leadtools.WinForms.RasterViewerInteractiveMode.Pan;

                // グレースケールに変換
                Leadtools.ImageProcessing.GrayscaleCommand grayScaleCommand = new Leadtools.ImageProcessing.GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                cs.Dispose();
                Leadtools.Codecs.RasterCodecs.Shutdown();
                //global.pblImagePath = tempImgName;
            }
            else
            {
                //画像ファイルがないとき
                lblNoImage.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;

                leadImg.Visible = true;
                button1.Enabled = false;
                //global.pblImagePath = string.Empty;
            }
        }

        private void leadImg_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }

        private void leadImg_MouseMove(object sender, MouseEventArgs e)
        {
            this.Cursor = Cursors.Hand;
        }
        
        /// -----------------------------------------------------------------------------
        /// <summary>
        ///     設定月数分経過した過去画像を削除する</summary>
        /// <param name="_dYYMM">
        ///     基準年月 (例：201401)</param>
        /// -----------------------------------------------------------------------------
        private void deleteImageArchived(int _dYYMM)
        {
            int _DataYYMM;
            string fileYYMM;

            // 設定月数分経過した過去画像を削除する            
            foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.tifPath, "*.tif"))
            {
                // ファイル名が規定外のファイルは読み飛ばします
                if (System.IO.Path.GetFileName(files).Length < 21)
                {
                    continue;
                }

                //ファイル名より年月を取得する
                fileYYMM = System.IO.Path.GetFileName(files).Substring(0, 6);

                _DataYYMM = Utility.StrtoInt(fileYYMM);

                if (_DataYYMM != global.flgOff)
                {
                    //基準年月以前なら削除する
                    if (_DataYYMM <= _dYYMM)
                    {
                        File.Delete(files);
                    }
                }
            }
        }

        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void txtYear_TextChanged(object sender, EventArgs e)
        {
            //// 曜日
            //DateTime eDate;
            //int tYY = Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei;
            //string sDate = tYY.ToString() + "/" + Utility.EmptytoZero(txtMonth.Text) + "/" +
            //        Utility.EmptytoZero(txtDay.Text);

            //// 存在する日付と認識された場合、曜日を表示する
            //if (DateTime.TryParse(sDate, out eDate))
            //{
            //    txtWeekDay.Text = ("日月火水木金土").Substring(int.Parse(eDate.DayOfWeek.ToString("d")), 1);
            //}
            //else
            //{
            //    txtWeekDay.Text = string.Empty;
            //}
        }
        
        private void dGV_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            //if (editLogStatus)
            //{
            //    if (e.ColumnIndex == 2 || e.ColumnIndex == 3 || e.ColumnIndex == 4 || e.ColumnIndex == 6 ||
            //        e.ColumnIndex == 7 || e.ColumnIndex == 9 || e.ColumnIndex == 10 || e.ColumnIndex == 12 ||
            //        e.ColumnIndex == 13 || e.ColumnIndex == 15 || e.ColumnIndex == 16 || e.ColumnIndex == 17)
            //    {
            //        dGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
            //        cellAfterValue = Utility.NulltoStr(dGV[e.ColumnIndex, e.RowIndex].Value);

            //        // 変更のとき編集ログデータを書き込み
            //        if (cellBeforeValue != cellAfterValue)
            //        {
            //            logDataUpdate(e.RowIndex, cI, global.flgOn, dGV.Columns[e.ColumnIndex].Name);
            //            dGV[e.ColumnIndex, e.RowIndex].Style.BackColor = Color.Cornsilk;
            //        }
            //    }
            //}
        }
        
        private void txtYear_Enter(object sender, EventArgs e)
        {

        }

        private void txtYear_Leave(object sender, EventArgs e)
        {

        }

        private void txtCode_TextChanged(object sender, EventArgs e)
        {

        }


        ///------------------------------------------------------------
        /// <summary>
        ///     就業奉行時間表記設定取得（0：24時間表記 1：48時間表記） </summary>
        /// <returns>
        ///     0：24時間表記 1：48時間表記 -1：エラー</returns>
        ///------------------------------------------------------------
        private int getLaborSystemCodeFigure()
        {
            int cf = -1;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);

            // 奉行SQLServer接続
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);
            SqlDataReader dR = null;

            try
            {
                // 就業奉行時間表記設定取得（0：24時間表記 1：48時間表記）
                StringBuilder sb = new StringBuilder();
                sb.Clear();
                sb.Append("select DateTimeIndication from tbHR_CorpOperationConfig ");
                
                dR = sdCon.free_dsReader(sb.ToString());

                while (dR.Read())
                {
                    cf = Utility.StrtoInt(dR["DateTimeIndication"].ToString());
                }

                return cf;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return cf;
            }
            finally
            {
                if (!dR.IsClosed)
                {
                    dR.Close();
                }

                if (sdCon.Cn.State == ConnectionState.Open)
                {
                    sdCon.Close();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //印刷確認
            if (!leadImg.Visible)
            {
                return;
            }

            if (MessageBox.Show("この伝票画像を印刷します。よろしいですか？", "印刷確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            //画像印刷
            cPrint prn = new cPrint();
            prn.Image(leadImg);
        }
    }
}
