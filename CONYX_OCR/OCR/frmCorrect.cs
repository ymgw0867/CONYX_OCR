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
    public partial class frmCorrect : Form
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
        public frmCorrect(string dbName, string comName, string sID)
        {
            InitializeComponent();

            _dbName = dbName;       // データベース名
            _comName = comName;     // 会社名
            dID = sID;              // 処理モード

            // テーブルアダプターマネージャーに勤務票ヘッダ、明細テーブルアダプターを割り付ける
            adpMn.勤務票ヘッダTableAdapter = hAdp;
            adpMn.勤務票明細TableAdapter = iAdp;

            // 編集ログデータ
            gAdp.Fill(dtsCl.出勤簿編集ログ);

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
        CONYX_CLIDataSetTableAdapters.TableAdapterManager adpMn = new CONYX_CLIDataSetTableAdapters.TableAdapterManager();
        CONYX_CLIDataSetTableAdapters.勤務票ヘッダTableAdapter hAdp = new CONYX_CLIDataSetTableAdapters.勤務票ヘッダTableAdapter();
        CONYX_CLIDataSetTableAdapters.勤務票明細TableAdapter iAdp = new CONYX_CLIDataSetTableAdapters.勤務票明細TableAdapter();
        CONYX_CLIDataSetTableAdapters.出勤簿編集ログTableAdapter gAdp = new CONYX_CLIDataSetTableAdapters.出勤簿編集ログTableAdapter();


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

            // 勤務データ登録
            if (dID == string.Empty)
            {
                // CSVデータをMDBへ読み込みます
                GetCsvDataToMDB();

                // データセットへデータを読み込みます
                getDataSet();

                // データテーブル件数カウント
                if (dtsCl.勤務票ヘッダ.Count == 0)
                {
                    MessageBox.Show("対象となる出勤簿データがありません", "出勤簿データ登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    
                    //終了処理
                    Environment.Exit(0);
                }

                // キー配列作成
                keyArrayCreate();
            }

            // 就業奉行時間表記設定取得（0：24時間表記 1：48時間表記）
            dtIndi = getLaborSystemCodeFigure();

            //// キャプション
            //this.Text = "出勤簿データ表示";

            // グリッドビュー定義
            GridviewSet gs = new GridviewSet();
            gs.Setting_Shain(dGV);

            // 編集作業、過去データ表示の判断
            if (dID == string.Empty) // パラメータのヘッダIDがないときは編集作業
            {
                // 最初のレコードを表示
                cI = 0;
                showOcrData(cI);
            }

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


        ///-------------------------------------------------------------
        /// <summary>
        ///     キー配列作成 </summary>
        ///-------------------------------------------------------------
        private void keyArrayCreate()
        {
            int iX = 0;
            foreach (var t in dtsCl.勤務票ヘッダ.OrderBy(a => a.ID))
            {
                Array.Resize(ref cIdx, iX + 1);
                cIdx[iX] = t.ID;
                iX++;
            }
        }

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
                            c.ReadOnly = false;
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

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                // 勤務体系コード：英数字入力可能とする　2018/08/31
                if (dGV.CurrentCell.ColumnIndex == 2)
                {
                    //イベントハンドラが複数回追加されてしまうので最初に削除する
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);

                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress2);
                }                        
                else if (dGV.CurrentCell.ColumnIndex > 3) // 数字のみ入力可能とする
                {
                    //イベントハンドラが複数回追加されてしまうので最初に削除する
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);

                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }
            }
        }

        void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        void Control_KeyPress2(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (e.KeyChar >= 'A' && e.KeyChar <= 'Z') ||
                e.KeyChar == '\b' || e.KeyChar == '\t')
                e.Handled = false;
            else e.Handled = true;
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!global.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            string colName = dGV.Columns[e.ColumnIndex].Name;

            // 過去データ表示のときは終了
            if (dID != string.Empty)
            {
                return;
            }

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
                {
                    dGV.Rows[e.RowIndex].DefaultCellStyle.BackColor = SystemColors.Control;
                }
                else
                {
                    dGV[cDay, e.RowIndex].Style.BackColor = SystemColors.Control;
                    dGV[cWeek, e.RowIndex].Style.BackColor = SystemColors.Control;
                    //dGV.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Empty;
                    dGV.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255, 254, 255);
                }
            }

            //// 社員番号のとき社員名を表示します
            //if (colName == cShainNum)
            //{
            //    // ChangeValueイベントを発生させない
            //    global.ChangeValueStatus = false;

            //    // 氏名を初期化
            //    dGV[cName, e.RowIndex].Value = string.Empty;
                
            //    // 奉行データベースより社員名を取得して表示します
            //    if (Utility.NulltoStr(dGV[cShainNum, e.RowIndex].Value) != string.Empty)
            //    {
            //        dbControl.DataControl dCon = new dbControl.DataControl(_dbName);
            //        string sYY = (Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei).ToString();
            //        string sMM = Utility.StrtoInt(txtMonth.Text).ToString();
            //        string sDD = Utility.StrtoInt(txtDay.Text).ToString();
            //        string sNum = Utility.NulltoStr(dGV[cShainNum, e.RowIndex].Value);
            //        OleDbDataReader dR = dCon.GetEmployeeBase(sYY, sMM, sDD, sNum);
            //        while (dR.Read())
            //        {
            //            // 所属名・社員名表示
            //            lblShozoku.Text = dR["DepartmentName"].ToString().Trim();
            //            dGV[cName, e.RowIndex].Value = dR["Name"].ToString().Trim();
            //            //dGV[cSzCode, e.RowIndex].Value = dR["DepartmentCode"].ToString().Trim().Substring(10, 5);

            //            // 2018/04/02
            //            if (dR["DepartmentCode"].ToString().Trim().Length > 14)
            //            {
            //                dGV[cSzCode, e.RowIndex].Value = dR["DepartmentCode"].ToString().Trim().Substring(10, 5);
            //            }
            //            else if (dR["DepartmentCode"].ToString().Trim().Length > 4)
            //            {
            //                dGV[cSzCode, e.RowIndex].Value = dR["DepartmentCode"].ToString().Trim().Substring(0, 5);
            //            }
            //            else
            //            {
            //                dGV[cSzCode, e.RowIndex].Value = dR["DepartmentCode"].ToString().Trim();
            //            }
            //        }

            //        dR.Close();
            //        dCon.Close();

            //        // 時刻区切り文字
            //        dGV[cSE, e.RowIndex].Value = ":";
            //        dGV[cEE, e.RowIndex].Value = ":";
            //        dGV[cSIE, e.RowIndex].Value = ":";
            //        dGV[cZE, e.RowIndex].Value = ":";
            //    }

            //    // ChangeValueイベントステータスをtrueに戻す
            //    global.ChangeValueStatus = true;
            //}
        }
        
        private void frmCorrect_Shown(object sender, EventArgs e)
        {
            if (dID != string.Empty) btnRtn.Focus();
        }

        private void dataGridView3_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                //イベントハンドラを追加する
                e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            }
        }

        private void dataGridView4_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                //イベントハンドラを追加する
                e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cIdx[cI]);

            //レコードの移動
            if (cI + 1 < dtsCl.勤務票ヘッダ.Rows.Count)
            {
                cI++;
                showOcrData(cI);
            }   
        }

        ///-----------------------------------------------------------------------------------
        /// <summary>
        ///     カレントデータを更新する</summary>
        /// <param name="iX">
        ///     カレントレコードのインデックス</param>
        ///-----------------------------------------------------------------------------------
        private void CurDataUpDate(string sID)
        {
            // エラーメッセージ
            string errMsg = "出勤簿テーブル更新";
            DateTime dt = DateTime.Today;

            try
            {
                // 勤務票ヘッダテーブル行を取得
                CONYX_CLIDataSet.勤務票ヘッダRow r = dtsCl.勤務票ヘッダ.Single(a => a.ID == sID);

                // 勤務票ヘッダテーブルセット更新
                r.年 = Utility.StrtoInt(txtYear.Text);
                r.月 = Utility.StrtoInt(txtMonth.Text);
                r.社員番号 = Utility.StrtoInt(txtCode.Text);
                r.社員名 = lblName.Text;
                r.確認 = Convert.ToInt32(chkKakunin.Checked);
                r.備考 = txtMemo.Text;
                r.更新年月日 = DateTime.Now;

                // 勤務票明細テーブルセット更新
                for (int i = 0; i < dGV.Rows.Count; i++)
                {
                    CONYX_CLIDataSet.勤務票明細Row m = (CONYX_CLIDataSet.勤務票明細Row)dtsCl.勤務票明細.FindByID(int.Parse(dGV[cID, i].Value.ToString()));

                    //m.日 = Utility.NulltoStr(dGV[cDay, i].Value);
                    m.勤務体系コード = Utility.NulltoStr(dGV[cKinmuTaikei, i].Value);
                    m.残業休日出勤申請 = Convert.ToInt32(dGV[cFlg, i].Value);
                    m.開始時 = timeValH(dGV[cSH, i].Value);
                    m.開始分 = timeVal(dGV[cSM, i].Value, 2);
                    m.退勤時 = timeValH(dGV[cEH, i].Value);
                    m.退勤分 = timeVal(dGV[cEM, i].Value, 2);
                    m.休憩時 = timeValH(dGV[cRH, i].Value);
                    m.休憩分 = timeVal(dGV[cRM, i].Value, 2);
                    m.実労時 = timeValH(dGV[cWH, i].Value);
                    m.実労分 = timeVal(dGV[cWM, i].Value, 2);
                    m.事由１ = Utility.NulltoStr(dGV[cJ1, i].Value);
                    m.事由2 = Utility.NulltoStr(dGV[cJ2, i].Value);
                    m.更新年月日 = DateTime.Now;
                    m.歴日付 = getCalenderDate(r.年, r.月, m.日);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, errMsg, MessageBoxButtons.OK);
            }
            finally
            {
            }
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

            // 後半(1日～15日）を当月とするため　コメント化 : 2018/08/31
            //if (sDD > 15)
            //{
            //    if (DateTime.TryParse(sYY + "/" + sMM + "/" + sDD, out dt))
            //    {
            //        rtnDate = dt;
            //    }
            //}
            //else if ((sMM + 1) > 12)
            //{
            //    if (DateTime.TryParse((sYY + 1) + "/01/" + sDD, out dt))
            //    {
            //        rtnDate = dt;
            //    }
            //}
            //else if (DateTime.TryParse(sYY + "/" + (sMM + 1) + "/" + sDD, out dt))
            //{
            //    rtnDate = dt;
            //}


            // 後半(1日～15日）を当月とする : 2018/08/31
            // 15日以降は暦で前月：2018/08/31
            if (sDD > 15)
            {
                // 前月16日～末日
                // 当月が1月の場合
                if (sMM == 1)
                {
                    // 前年12月
                    if (DateTime.TryParse((sYY - 1) + "/12/" + sDD, out dt))
                    {
                        rtnDate = dt;
                    }
                }
                else if (DateTime.TryParse(sYY + "/" + (sMM - 1) + "/" + sDD, out dt))
                {
                    rtnDate = dt;
                }
            }
            else if (DateTime.TryParse(sYY + "/" + sMM + "/" + sDD, out dt))
            {
                // 当月1日～15日
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
            //カレントデータの更新
            CurDataUpDate(cIdx[cI]);

            //レコードの移動
            cI =  dtsCl.勤務票ヘッダ.Rows.Count - 1;
            showOcrData(cI);
        }

        private void btnBefore_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cIdx[cI]);

            //レコードの移動
            if (cI > 0)
            {
                cI--;
                showOcrData(cI);
            }   
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cIdx[cI]);

            //レコードの移動
            cI = 0;
            showOcrData(cI);
        }

        /// <summary>
        /// エラーチェックボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnErrCheck_Click(object sender, EventArgs e)
        {
            // 非ログ書き込み状態とする：2015/09/25
            editLogStatus = false;

            // OCRDataクラス生成
            OCRData ocr = new OCRData(_dbName);

            // エラーチェックを実行
            if (getErrData(cI, ocr))
            {
                MessageBox.Show("エラーはありませんでした", "エラーチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dGV.CurrentCell = null;

                // データ表示
                showOcrData(cI);
            }
            else
            {
                // カレントインデックスをエラーありインデックスで更新
                cI = ocr._errHeaderIndex;

                // データ表示
                showOcrData(cI);

                // エラー表示
                ErrShow(ocr);
            }
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cIdx[cI]);

            //レコードの移動
            cI = hScrollBar1.Value;
            showOcrData(cI);
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の出勤簿データを削除します。よろしいですか", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            // 非ログ書き込み状態とする
            editLogStatus = false;

            // レコードと画像ファイルを削除する
            DataDelete(cI);

            // 勤務票ヘッダテーブル件数カウント
            if (dtsCl.勤務票ヘッダ.Count() > 0)
            {
                // カレントレコードインデックスを再設定
                if (dtsCl.勤務票ヘッダ.Count() - 1 < cI)
                {
                    cI = dtsCl.勤務票ヘッダ.Count() - 1;
                }

                // データ画面表示
                showOcrData(cI);
            }
            else
            {
                // ゼロならばプログラム終了
                MessageBox.Show("全ての出勤簿データが削除されました。処理を終了します。", "勤務票削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                //終了処理
                this.Tag = END_NODATA;
                this.Close();
            }
        }

        ///-------------------------------------------------------------------------------
        /// <summary>
        ///     １．指定した勤務票ヘッダデータと勤務票明細データを削除する　
        ///     ２．該当する画像データを削除する</summary>
        /// <param name="i">
        ///     勤務票ヘッダRow インデックス</param>
        ///-------------------------------------------------------------------------------
        private void DataDelete(int i)
        {
            string sImgNm = string.Empty;
            string errMsg = string.Empty;

            // 勤務票データ削除
            try
            {
                // 出勤簿編集ログデータベース更新
                gAdp.Update(dtsCl.出勤簿編集ログ);
                gAdp.Fill(dtsCl.出勤簿編集ログ);

                // ヘッダIDを取得します
                CONYX_CLIDataSet.勤務票ヘッダRow r = dtsCl.勤務票ヘッダ.Single(a => a.ID == cIdx[i]);

                // 画像ファイル名を取得します
                sImgNm = r.画像名;

                // データテーブルからヘッダIDが一致する勤務票明細データを削除します。
                errMsg = "勤務票明細データ";
                foreach (CONYX_CLIDataSet.勤務票明細Row item in dtsCl.勤務票明細.Rows)
                {
                    if (item.RowState != DataRowState.Deleted && item.ヘッダID == r.ID)
                    {
                        item.Delete();
                    }
                }

                // 出勤簿編集ログ削除
                errMsg = "出勤簿編集ログ";
                for (int iX = 0; iX < dtsCl.出勤簿編集ログ.Rows.Count; iX++)
                {
                    CONYX_CLIDataSet.出勤簿編集ログRow lr = (CONYX_CLIDataSet.出勤簿編集ログRow)dtsCl.出勤簿編集ログ.Rows[iX];

                    if (lr.RowState != DataRowState.Deleted && lr.勤務表ヘッダID == r.ID)
                    {
                        dtsCl.出勤簿編集ログ.Rows[iX].Delete();
                    }
                }

                //dtsCl.出勤簿編集ログ.AcceptChanges();

                // データテーブルから勤務票ヘッダデータを削除します
                errMsg = "勤務票ヘッダデータ";
                r.Delete();

                // データベース更新
                adpMn.UpdateAll(dtsCl);
                gAdp.Update(dtsCl.出勤簿編集ログ);

                // 画像ファイルを削除します
                errMsg = "勤務管理表画像";
                if (sImgNm != string.Empty)
                {
                    if (System.IO.File.Exists(Properties.Settings.Default.dataPath + sImgNm))
                    {
                        System.IO.File.Delete(Properties.Settings.Default.dataPath + sImgNm);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(errMsg + "の削除に失敗しました" + Environment.NewLine + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                //配列キー再構築
                keyArrayCreate();
            }

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

                // カレントデータ更新
                if (dID == string.Empty)
                {
                    CurDataUpDate(cIdx[cI]);
                }
            }

            // データベース更新
            adpMn.UpdateAll(dtsCl);
            gAdp.Update(dtsCl.出勤簿編集ログ);

            // 解放する
            this.Dispose();
        }

        private void btnDataMake_Click(object sender, EventArgs e)
        {
            // 非ログ書き込み状態とする
            editLogStatus = false;

            // 就業奉行用CSVデータ出力
            textDataMake();
        }

        /// -----------------------------------------------------------------------
        /// <summary>
        ///     就業奉行・受入CSVデータ出力 </summary>
        /// -----------------------------------------------------------------------
        private void textDataMake()
        {
            if (MessageBox.Show("就業奉行受け渡しデータを作成します。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            // OCRDataクラス生成
            OCRData ocr = new OCRData(_dbName);

            // エラーチェックを実行する
            if (getErrData(cI, ocr)) // エラーがなかったとき
            {
                // データベース更新
                adpMn.UpdateAll(dtsCl);
                gAdp.Update(dtsCl.出勤簿編集ログ);
                gAdp.Fill(dtsCl.出勤簿編集ログ);

                // OCROutputクラス インスタンス生成
                OCROutput kd = new OCROutput(this, dtsCl, dtIndi);

                // 汎用データ作成
                kd.SaveData(_dbName);          

                // 画像ファイル退避
                tifFileMove();

                // データベース更新
                adpMn.UpdateAll(dtsCl);

                // 過去データ作成
                saveLastData();

                // 設定月数分経過した過去画像と過去勤務データを削除する
                deleteArchived();

                // 勤務票データ削除
                deleteDataAll();

                // MDBファイル最適化
                mdbCompact();

                //終了
                MessageBox.Show("終了しました。就業奉行ｉで勤務データ受け入れを行ってください。", "就業奉行受け入れデータ作成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Tag = END_MAKEDATA;
                this.Close();
            }
            else
            {
                // カレントインデックスをエラーありインデックスで更新
                cI = ocr._errHeaderIndex;

                // エラーあり
                showOcrData(cI);    // データ表示
                ErrShow(ocr);   // エラー表示
            }
        }

        /// -----------------------------------------------------------------------------------
        /// <summary>
        ///     エラーチェックを実行する</summary>
        /// <param name="cIdx">
        ///     現在表示中の勤務票ヘッダデータインデックス</param>
        /// <param name="ocr">
        ///     OCRDATAクラスインスタンス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        /// -----------------------------------------------------------------------------------
        private bool getErrData(int cId, OCRData ocr)
        {
            // カレントレコード更新
            CurDataUpDate(cIdx[cId]);

            // エラー番号初期化
            ocr._errNumber = ocr.eNothing;

            // エラーメッセージクリーン
            ocr._errMsg = string.Empty;

            // エラーチェック実行①:カレントレコードから最終レコードまで
            if (!ocr.errCheckMain(cId, (dtsCl.勤務票ヘッダ.Rows.Count - 1), this, dtsCl, cIdx))
            {
                return false;
            }

            // エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (cId > 0)
            {
                if (!ocr.errCheckMain(0, (cId - 1), this, dtsCl, cIdx))
                {
                    return false;
                }
            }

            // エラーなし
            lblErrMsg.Text = string.Empty;

            return true;
        }
        
        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     画像ファイル退避処理 </summary>
        ///----------------------------------------------------------------------------------
        private void tifFileMove()
        {
            // 移動先フォルダ名
            string dirName = Properties.Settings.Default.tifPath + global.cnfYear + global.cnfMonth.ToString().PadLeft(2, '0') + @"\";

            // 移動先フォルダがあるか？なければ作成する（TIFフォルダ内の年月フォルダ）
            if (!System.IO.Directory.Exists(dirName))
            {
                System.IO.Directory.CreateDirectory(dirName);
            }

            // 出勤簿ヘッダデータを取得する
            var s = dtsCl.勤務票ヘッダ.OrderBy(a => a.ID);

            int nn = 0;

            foreach (var t in s)
            {
                string NewFilenameYearMonth = t.年.ToString() + t.月.ToString().PadLeft(2, '0');

                // 画像ファイルパスを取得する
                string fromImg = Properties.Settings.Default.dataPath + t.画像名;

                // ファイル名を「対象年月_社員番号_社員名」に変えて退避先フォルダへ移動する
                string NewFilename = NewFilenameYearMonth + "_" + t.社員番号.ToString().PadLeft(8, '0') +  "_" + t.社員名;

                // 移動先ファイルパス
                string toImg = string.Empty;

                // 同名ファイルが既に登録済みのときは枝番を付加する
                // 同名ファイルが既に登録済みのときタイムスタンプを付加する
                int fCnt = System.IO.Directory.GetFiles(dirName, NewFilename + "*.tif").Count();
                if (fCnt > 0)
                {
                    toImg = dirName + NewFilename + "_" + DateTime.Now.Hour.ToString().PadLeft(2, '0') + DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0') + nn + ".tif";
                }
                else
                {
                    toImg = dirName + NewFilename + ".tif";
                }

                // ファイルを移動する
                if (System.IO.File.Exists(fromImg))
                {
                    System.IO.File.Move(fromImg, toImg);
                }

                // 出勤簿ヘッダレコードの画像ファイル名を書き換える
                t.画像名 = System.IO.Path.GetFileName(toImg);

                nn++;
            }
            
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

        /// ---------------------------------------------------------------------------------
        /// <summary>
        ///     設定月数分経過した過去画像と過去勤務データを削除する </summary> 
        /// ---------------------------------------------------------------------------------
        private void deleteArchived()
        {
            // 削除月設定が0のとき、「過去画像削除しない」とみなし終了する
            if (global.cnfArchived == global.flgOff)
            {
                return;
            }

            try
            {
                // 削除年月の取得
                DateTime dt = DateTime.Parse(DateTime.Today.Year.ToString() + "/" + DateTime.Today.Month.ToString() + "/01");
                DateTime delDate = dt.AddMonths(global.cnfArchived * (-1));
                int _dYY = delDate.Year;            //基準年
                int _dMM = delDate.Month;           //基準月
                int _dYYMM = _dYY * 100 + _dMM;     //基準年月
                //int _waYYMM = (delDate.Year - Properties.Settings.Default.rekiHosei) * 100 + _dMM;   //基準年月(和暦）

                //// 設定月数分経過した過去画像を削除する
                //deleteImageArchived(_dYYMM);

                // 設定月数分経過した過去画像を削除する
                // 過去勤務票データを削除する
                deleteLastDataArchived(_dYYMM);
            }
            catch (Exception e)
            {
                MessageBox.Show("過去画像・過去勤務票データ削除中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
                return;
            }
            finally
            {
                //if (ocr.sCom.Connection.State == ConnectionState.Open) ocr.sCom.Connection.Close();
            }
        }

        /// ---------------------------------------------------------------------------
        /// <summary>
        ///     過去勤務票データ登録 </summary>
        /// ---------------------------------------------------------------------------
        private void saveLastData()
        {
            try
            {
                ////  過去勤務票ヘッダデータとその明細データを削除します
                //deleteLastData();

                // データセットへデータを再読み込みします
                getDataSet();

                // 過去勤務票ヘッダデータと過去勤務票明細データを作成します
                addLastdata();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "過去勤務票データ作成エラー", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     過去勤務票ヘッダデータとその明細データを削除します</summary>    
        ///     
        /// -------------------------------------------------------------------------
        private void deleteLastData()
        {
            //OleDbCommand sCom = new OleDbCommand();
            //OleDbCommand sCom2 = new OleDbCommand();
            //OleDbCommand sCom3 = new OleDbCommand();
            //mdbControl mdb = new mdbControl();
            //mdb.dbConnectCl(sCom);
            //mdb.dbConnect(sCom2);
            //mdb.dbConnect(sCom3);
            //OleDbDataReader dR = null;
            //OleDbDataReader dR2 = null;

            //StringBuilder sb = new StringBuilder();
            //StringBuilder sbd = new StringBuilder();

            //try
            //{
            //    // 対象データ : 取消は対象外とする　2015/10/01
            //    sb.Clear();
            //    sb.Append("Select 勤務票明細.ヘッダID, 勤務票明細.ID,");
            //    sb.Append("勤務票ヘッダ.年, 勤務票ヘッダ.月, 勤務票ヘッダ.日,");
            //    sb.Append("勤務票明細.社員番号 from 勤務票ヘッダ inner join 勤務票明細 ");
            //    sb.Append("on 勤務票ヘッダ.ID = 勤務票明細.ヘッダID ");
            //    sb.Append("where 勤務票明細.取消 = 0 ");
            //    sb.Append("order by 勤務票明細.ヘッダID, 勤務票明細.ID");

            //    sCom.CommandText = sb.ToString();
            //    dR = sCom.ExecuteReader();

            //    while (dR.Read())
            //    {
            //        // ヘッダID
            //        string hdID = string.Empty;

            //        // 日付と社員番号で過去データを抽出（該当するのは1件）
            //        sb.Clear();
            //        sb.Append("Select 過去勤務票明細.ヘッダID,過去勤務票明細.ID,");
            //        sb.Append("過去勤務票ヘッダ.年, 過去勤務票ヘッダ.月, 過去勤務票ヘッダ.日,");
            //        sb.Append("過去勤務票明細.社員番号 from 過去勤務票ヘッダ inner join 過去勤務票明細 ");
            //        sb.Append("on 過去勤務票ヘッダ.ID = 過去勤務票明細.ヘッダID ");
            //        sb.Append("where ");
            //        sb.Append("過去勤務票ヘッダ.年 = ? and ");
            //        sb.Append("過去勤務票ヘッダ.月 = ? and ");
            //        sb.Append("過去勤務票ヘッダ.日 = ? and ");
            //        sb.Append("過去勤務票ヘッダ.データ領域名 = ? and ");
            //        sb.Append("過去勤務票明細.社員番号 = ?");

            //        sCom2.CommandText = sb.ToString();
            //        sCom2.Parameters.Clear();
            //        sCom2.Parameters.AddWithValue("@yy", dR["年"].ToString());
            //        sCom2.Parameters.AddWithValue("@mm", dR["月"].ToString());
            //        sCom2.Parameters.AddWithValue("@dd", dR["日"].ToString());
            //        sCom2.Parameters.AddWithValue("@db", _dbName);
            //        sCom2.Parameters.AddWithValue("@n", dR["社員番号"].ToString());

            //        dR2 = sCom2.ExecuteReader();

            //        while (dR2.Read())
            //        {
            //            // ヘッダIDを取得
            //            if (hdID == string.Empty)
            //            {
            //                hdID = dR2["ヘッダID"].ToString();
            //            }

            //            // 過去勤務票明細レコード削除
            //            sbd.Clear();
            //            sbd.Append("delete from 過去勤務票明細 ");
            //            sbd.Append("where ID = ?");

            //            sCom3.CommandText = sbd.ToString();
            //            sCom3.Parameters.Clear();
            //            sCom3.Parameters.AddWithValue("@id", dR2["ID"].ToString());

            //            sCom3.ExecuteNonQuery();
            //        }

            //        dR2.Close();
            //    }

            //    dR.Close();

            //    // データベース接続解除
            //    if (sCom.Connection.State == ConnectionState.Open)
            //    {
            //        sCom.Connection.Close();
            //    }

            //    if (sCom2.Connection.State == ConnectionState.Open)
            //    {
            //        sCom2.Connection.Close();
            //    }

            //    if (sCom3.Connection.State == ConnectionState.Open)
            //    {
            //        sCom3.Connection.Close();
            //    }

            //    // データベース再接続
            //    mdb.dbConnect(sCom);
            //    mdb.dbConnect(sCom2);

            //    // 明細データのない過去勤務票ヘッダデータを抽出
            //    sb.Clear();
            //    sb.Append("Select 過去勤務票ヘッダ.ID,過去勤務票明細.ヘッダID ");
            //    sb.Append("from 過去勤務票ヘッダ left join 過去勤務票明細 ");
            //    sb.Append("on 過去勤務票ヘッダ.ID = 過去勤務票明細.ヘッダID ");
            //    sb.Append("where ");
            //    sb.Append("過去勤務票明細.ヘッダID is null");
            //    sCom.CommandText = sb.ToString();
            //    dR = sCom.ExecuteReader();

            //    while (dR.Read())
            //    {
            //        // 過去勤務票ヘッダレコード削除
            //        sbd.Clear();

            //        sbd.Append("delete from 過去勤務票ヘッダ ");
            //        sbd.Append("where ID = ?");

            //        sCom2.CommandText = sbd.ToString();
            //        sCom2.Parameters.Clear();
            //        sCom2.Parameters.AddWithValue("@id", dR["ID"].ToString());

            //        sCom2.ExecuteNonQuery();
            //    }

            //    dR.Close();
            //}
            //catch (Exception e)
            //{
            //    MessageBox.Show(e.Message);
            //}
            //finally
            //{
            //    if (sCom.Connection.State == ConnectionState.Open)
            //    {
            //        sCom.Connection.Close();
            //    }

            //    if (sCom2.Connection.State == ConnectionState.Open)
            //    {
            //        sCom2.Connection.Close();
            //    }

            //    if (sCom3.Connection.State == ConnectionState.Open)
            //    {
            //        sCom3.Connection.Close();
            //    }
            //}
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     過去勤務票ヘッダデータと過去勤務票明細データを作成します</summary>
        ///     
        /// -------------------------------------------------------------------------
        private void addLastdata()
        {
            // 過去勤務票ヘッダテーブルデータ読み込み
            CONYXDataSetTableAdapters.過去勤務票ヘッダTableAdapter lsh = new CONYXDataSetTableAdapters.過去勤務票ヘッダTableAdapter();
            //lsh.Fill(dts.過去勤務票ヘッダ);

            // 過去勤務票明細テーブルデータ読み込み
            CONYXDataSetTableAdapters.過去勤務票明細TableAdapter lsm = new CONYXDataSetTableAdapters.過去勤務票明細TableAdapter();
            //lsm.Fill(dts.過去勤務票明細);

            for (int i = 0; i < dtsCl.勤務票ヘッダ.Rows.Count; i++)
            {
                // -------------------------------------------------------------------------
                //      過去勤務票ヘッダレコードを作成します
                // -------------------------------------------------------------------------
                CONYX_CLIDataSet.勤務票ヘッダRow hr = (CONYX_CLIDataSet.勤務票ヘッダRow)dtsCl.勤務票ヘッダ.Rows[i];
                CONYXDataSet.過去勤務票ヘッダRow nr = dts.過去勤務票ヘッダ.New過去勤務票ヘッダRow();

                #region テーブルカラム名比較～データコピー

                // 勤務票ヘッダのカラムを順番に読む
                for (int j = 0; j < dtsCl.勤務票ヘッダ.Columns.Count; j++)
                {
                    // 過去勤務票ヘッダのカラムを順番に読む
                    for (int k = 0; k < dts.過去勤務票ヘッダ.Columns.Count; k++)
                    {
                        // フィールド名が同じであること
                        if (dtsCl.勤務票ヘッダ.Columns[j].ColumnName == dts.過去勤務票ヘッダ.Columns[k].ColumnName)
                        {
                            if (dts.過去勤務票ヘッダ.Columns[k].ColumnName == "更新年月日")
                            {
                                nr[k] = DateTime.Now;   // 更新年月日はこの時点のタイムスタンプを登録
                            }
                            else
                            {
                                nr[k] = hr[j];          // データをコピー
                            }
                            break;
                        }
                    }
                }
                #endregion

                // 過去勤務票ヘッダデータテーブルに追加
                dts.過去勤務票ヘッダ.Add過去勤務票ヘッダRow(nr);

                // -------------------------------------------------------------------------
                //      過去勤務票明細レコードを作成します
                // -------------------------------------------------------------------------
                var mm = dtsCl.勤務票明細
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                           a.ヘッダID == hr.ID)
                    .OrderBy(a => a.ID);

                foreach (var item in mm)
                {
                    //// 非該当日付は対象外
                    //if (item.歴日付 == global.NODATE)
                    //{
                    //    continue;
                    //}
                    
                    CONYXDataSet.過去勤務票明細Row nm = dts.過去勤務票明細.New過去勤務票明細Row();

                    #region  テーブルカラム名比較～データコピー

                    // 勤務票明細のカラムを順番に読む
                    for (int j = 0; j < dtsCl.勤務票明細.Columns.Count; j++)
                    {
                        // IDはオートナンバーのため値はコピーしない
                        if (dtsCl.勤務票明細.Columns[j].ColumnName != "ID")
                        {
                            // 過去勤務票ヘッダのカラムを順番に読む
                            for (int k = 0; k < dts.過去勤務票明細.Columns.Count; k++)
                            {
                                // フィールド名が同じであること
                                if (dtsCl.勤務票明細.Columns[j].ColumnName == dts.過去勤務票明細.Columns[k].ColumnName)
                                {
                                    if (dts.過去勤務票明細.Columns[k].ColumnName == "更新年月日")
                                    {
                                        nm[k] = DateTime.Now;   // 更新年月日はこの時点のタイムスタンプを登録
                                    }
                                    else
                                    {
                                        nm[k] = item[j];          // データをコピー
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    #endregion

                    // 過去勤務票明細データテーブルに追加
                    dts.過去勤務票明細.Add過去勤務票明細Row(nm);
                }
            }

            // データベース更新
            lsh.Update(dts);
            lsm.Update(dts);

            // ローカルの出勤簿編集ログを共通の出勤簿編集ログに書き込み
            CONYXDataSetTableAdapters.出勤簿編集ログTableAdapter lAdp = new CONYXDataSetTableAdapters.出勤簿編集ログTableAdapter();

            for (int i = 0; i < dtsCl.出勤簿編集ログ.Rows.Count; i++)
            {
                CONYX_CLIDataSet.出勤簿編集ログRow r = (CONYX_CLIDataSet.出勤簿編集ログRow)dtsCl.出勤簿編集ログ.Rows[i];
                CONYXDataSet.出勤簿編集ログRow nr = dts.出勤簿編集ログ.New出勤簿編集ログRow();

                // ローカルの出勤簿編集ログのカラムを順番に読む
                for (int j = 0; j < dtsCl.出勤簿編集ログ.Columns.Count; j++)
                {
                    // 共通の出勤簿編集ログのカラムを順番に読む
                    for (int k = 0; k < dts.出勤簿編集ログ.Columns.Count; k++)
                    {
                        if (dtsCl.出勤簿編集ログ.Columns[j].ColumnName == "ID")
                        {
                            continue;
                        }

                        // フィールド名が同じであること
                        if (dtsCl.出勤簿編集ログ.Columns[j].ColumnName == dts.出勤簿編集ログ.Columns[k].ColumnName)
                        {
                            // データをコピー
                            nr[k] = r[j];

                            break;
                        }
                    }
                }

                // 共通の出勤簿編集ログテーブルに追加
                dts.出勤簿編集ログ.Add出勤簿編集ログRow(nr);
            }

            // データベース更新
            lAdp.Update(dts.出勤簿編集ログ);
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
                lblNoImage.Visible = false;
                global.pblImagePath = string.Empty;
                return;
            }

            //画像ファイルがあるとき表示
            if (File.Exists(tempImgName))
            {
                lblNoImage.Visible = false;
                leadImg.Visible = true;

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

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     基準年月以前の過去勤務票ヘッダデータとその明細データを削除します</summary>
        /// <param name="sYYMM">
        ///     基準年月</param>     
        /// -------------------------------------------------------------------------
        private void deleteLastDataArchived(int sYYMM)
        {
            CONYXDataSetTableAdapters.過去勤務票ヘッダTableAdapter adpH = new CONYXDataSetTableAdapters.過去勤務票ヘッダTableAdapter();
            CONYXDataSetTableAdapters.過去勤務票明細TableAdapter adpM = new CONYXDataSetTableAdapters.過去勤務票明細TableAdapter();

            adpH.Fill(dts.過去勤務票ヘッダ);
            adpM.Fill(dts.過去勤務票明細);

            // 基準年月以前の過去勤務票ヘッダデータを取得します
            var h = dts.過去勤務票ヘッダ
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                                a.年 * 100 + a.月 <= sYYMM);

            // foreach用の配列を作成
            var hLst = h.ToList();

            foreach (var lh in hLst)
            {
                // ヘッダIDが一致する過去勤務票明細を取得します
                var m = dts.過去勤務票明細
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                                a.ヘッダID == lh.ID);

                // foreach用の配列を作成
                var list = m.ToList();

                // 該当過去勤務票明細を削除します
                foreach (var lm in list)
                {
                    CONYXDataSet.過去勤務票明細Row lRow = (CONYXDataSet.過去勤務票明細Row)dts.過去勤務票明細.Rows.Find(lm.ID);
                    lRow.Delete();
                }

                // 画像ファイルを削除します
                string dirName = Properties.Settings.Default.tifPath + lh.画像名.Substring(0, 6) + @"\";
                string imgPath = dirName + lh.画像名;
                File.Delete(imgPath);

                // 過去勤務票ヘッダを削除します
                lh.Delete();
            }

            // データベース更新
            adpH.Update(dts);
            adpM.Update(dts);
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

        /// -------------------------------------------------------------------
        /// <summary>
        ///     勤務票ヘッダデータと勤務票明細データを全件削除します</summary>
        /// -------------------------------------------------------------------
        private void deleteDataAll() 
        {
            // 勤務票明細全行削除
            var m = dtsCl.勤務票明細.Where(a => a.RowState != DataRowState.Deleted);
            foreach (var t in m)
            {
                t.Delete();
            }

            // 勤務票ヘッダ全行削除
            var h = dtsCl.勤務票ヘッダ.Where(a => a.RowState != DataRowState.Deleted);
            foreach (var t in h)
            {
                t.Delete();
            }

            // データベース更新
            adpMn.UpdateAll(dtsCl);

            // ローカル出勤簿編集ログ全行削除
            var g = dtsCl.出勤簿編集ログ.Where(a => a.RowState != DataRowState.Deleted);
            foreach (var t in g)
            {
                t.Delete();
            }

            gAdp.Update(dtsCl.出勤簿編集ログ);

            // 後片付け
            dtsCl.勤務票明細.Dispose();
            dtsCl.勤務票ヘッダ.Dispose();
            dtsCl.出勤簿編集ログ.Dispose();
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
            if (editLogStatus)
            {
                if (e.ColumnIndex == 2 || e.ColumnIndex == 3 || e.ColumnIndex == 4 || e.ColumnIndex == 6 ||
                    e.ColumnIndex == 7 || e.ColumnIndex == 9 || e.ColumnIndex == 10 || e.ColumnIndex == 12 ||
                    e.ColumnIndex == 13 || e.ColumnIndex == 15 || e.ColumnIndex == 16 || e.ColumnIndex == 17)
                {
                    dGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    cellAfterValue = Utility.NulltoStr(dGV[e.ColumnIndex, e.RowIndex].Value);

                    // 変更のとき編集ログデータを書き込み
                    if (cellBeforeValue != cellAfterValue)
                    {
                        logDataUpdate(e.RowIndex, cI, global.flgOn, dGV.Columns[e.ColumnIndex].Name);

                        if (global.cnfEditBackColor != "")
                        {
                            dGV[e.ColumnIndex, e.RowIndex].Style.BackColor = Color.FromArgb(Convert.ToInt32(global.cnfEditBackColor, 16));
                        }
                    }
                }
            }
        }

        /// ----------------------------------------------------------------------
        /// <summary>
        ///     編集ログデータ書き込み </summary>
        /// <param name="rIndex">
        ///     データグリッドビュー行インデックス</param>
        /// ----------------------------------------------------------------------
        private void logDataUpdate(int rIndex, int iX, int dType, string colName)
        {
            CONYX_CLIDataSet.出勤簿編集ログRow r = dtsCl.出勤簿編集ログ.New出勤簿編集ログRow();

            // 勤務票ヘッダテーブル行から画像ファイル名を取得
            var h = dtsCl.勤務票ヘッダ.Single(a => a.ID == cIdx[iX]);

            r.勤務表ヘッダID = h.ID;
            r.社員番号 = h.社員番号.ToString();
            r.社員名 = h.社員名;
            r.年月日時刻 = DateTime.Now;
            r.編集日付 = getCalenderDate(h.年, h.月, Utility.StrtoInt(Utility.NulltoStr(dGV[cDay, rIndex].Value)));
            r.行番号 = rIndex;
            r.列名 = colName;

            // グリッド情報
            if (dType == global.flgOn)
            {
                // 社員の勤怠情報編集のとき
                r.行番号 = rIndex + 1;
            }
            else
            {
                // 出勤簿のヘッダー情報のとき
                r.行番号 = 0;
            }
            
            r.項目名 = cellName;
            r.変更前 = cellBeforeValue;
            r.変更後 = cellAfterValue;
            r.編集アカウント = global.flgOff;

            // 出勤簿編集ログデータを追加
            dtsCl.出勤簿編集ログ.Add出勤簿編集ログRow(r);
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            if (editLogStatus)
            {
                if (sender == txtYear) cellName = LOG_YEAR;
                if (sender == txtMonth) cellName = LOG_MONTH;
                if (sender == txtCode) cellName = LOG_NUMBER;

                TextBox tb = (TextBox)sender;

                // 値を保持
                cellBeforeValue = Utility.NulltoStr(tb.Text);
            }
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            if (editLogStatus)
            {
                TextBox tb = (TextBox)sender;
                cellAfterValue = Utility.NulltoStr(tb.Text);

                // 変更のとき編集ログデータを書き込み
                if (cellBeforeValue != cellAfterValue)
                {
                    logDataUpdate(0, cI, global.flgOff, string.Empty);

                    if (global.cnfEditBackColor != "")
                    {
                        tb.BackColor = Color.FromArgb(Convert.ToInt32(global.cnfEditBackColor, 16));
                    }
                }
            }
        }

        private void txtCode_TextChanged(object sender, EventArgs e)
        {
            // 氏名を初期化
            lblName.Text = string.Empty;

            // 奉行データベースより社員名を取得して表示します
            if (txtCode.Text != string.Empty)
            {
                // 接続文字列取得
                string sc = sqlControl.obcConnectSting.get(_dbName);
                sqlControl.DataControl sdCon = new common.sqlControl.DataControl(sc);

                string bCode = txtCode.Text.PadLeft(10, '0');
                SqlDataReader dR = sdCon.free_dsReader(Utility.getEmployee(bCode));

                while (dR.Read())
                {
                    // 社員名表示
                    if (Utility.StrtoInt(dR["zaisekikbn"].ToString()) == 2)
                    {
                        lblName.ForeColor = Color.Red;
                        lblName.Text = dR["Name"].ToString().Trim() + "（退）";
                    }
                    else
                    {
                        lblName.ForeColor = global.defaultColor;
                        lblName.Text = dR["Name"].ToString().Trim();
                    }
                }

                dR.Close();
                sdCon.Close();
            }
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

            if (MessageBox.Show("この出勤簿画像を印刷します。よろしいですか？", "印刷確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            //画像印刷
            cPrint prn = new cPrint();
            prn.Image(leadImg);
        }
    }
}
