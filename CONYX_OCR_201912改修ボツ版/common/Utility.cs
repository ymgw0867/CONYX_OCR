using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;
//using Excel = Microsoft.Office.Interop.Excel;

namespace CONYX_OCR.common
{
    class Utility
    {
        ///---------------------------------------------------------------------------
        /// <summary>
        ///     ウィンドウ最小サイズの設定 </summary>
        /// <param name="tempFrm">
        ///     対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">
        ///     width</param>
        /// <param name="hSize">
        ///     Height</param>
        ///---------------------------------------------------------------------------
        public static void WindowsMinSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MinimumSize = new Size(wSize, hSize);
        }

        ///---------------------------------------------------------------------------
        /// <summary>
        ///     ウィンドウ最小サイズの設定 </summary>
        /// <param name="tempFrm">
        ///     対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">
        ///     width</param>
        /// <param name="hSize">
        ///     height</param>
        ///---------------------------------------------------------------------------
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new Size(wSize, hSize);
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     ログインアカウントコンボボックスクラス </summary>
        ///--------------------------------------------------------
        //public class comboAccount
        //{
        //    public int code { get; set; }
        //    public string Name { get; set; }

        //    ///------------------------------------------------------------------------
        //    /// <summary>
        //    ///     ログインアカウントコンボボックスデータロード</summary>
        //    /// <param name="tempBox">
        //    ///     ロード先コンボボックスオブジェクト名</param>
        //    ///------------------------------------------------------------------------
        //    public static void Load(ComboBox tempBox, CBS_CLIDataSetTableAdapters.ログインユーザーTableAdapter uAdp)
        //    {
        //        CBSDataSet.ログインユーザーDataTable dt = uAdp.GetData();

        //        try
        //        {
        //            comboAccount cmb1;

        //            tempBox.Items.Clear();
        //            tempBox.DisplayMember = "Name";
        //            tempBox.ValueMember = "code";

        //            foreach (var a in dt.OrderBy(a => a.ID))
        //            {
        //                cmb1 = new comboAccount();
        //                cmb1.code = a.ID;
        //                cmb1.Name = a.名前;
        //                tempBox.Items.Add(cmb1);
        //            }

        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message, "アカウントコンボボックスロード");
        //        }
        //    }

        //    ///------------------------------------------------------------------------
        //    /// <summary>
        //    ///     ログインアカウントコンボ表示 </summary>
        //    /// <param name="tempBox">
        //    ///     コンボボックスオブジェクト</param>
        //    /// <param name="dt">
        //    ///     月日</param>
        //    ///------------------------------------------------------------------------
        //    public static void selectedIndex(ComboBox tempBox, int code)
        //    {
        //        comboAccount cmbS = new comboAccount();
        //        Boolean Sh = false;

        //        for (int iX = 0; iX <= tempBox.Items.Count - 1; iX++)
        //        {
        //            tempBox.SelectedIndex = iX;
        //            cmbS = (comboAccount)tempBox.SelectedItem;

        //            if (cmbS.code == code)
        //            {
        //                Sh = true;
        //                break;
        //            }
        //        }

        //        if (Sh == false)
        //        {
        //            tempBox.SelectedIndex = -1;
        //        }
        //    }
        //}

        ///--------------------------------------------------------
        /// <summary>
        ///     休日コンボボックスクラス </summary>
        ///--------------------------------------------------------
        public class comboHoliday
        {
            public string Date { get; set; }
            public string Name { get; set; }

            ///------------------------------------------------------------------------
            /// <summary>
            ///     休日コンボボックスデータロード</summary>
            /// <param name="tempBox">
            ///     ロード先コンボボックスオブジェクト名</param>
            ///------------------------------------------------------------------------
            public static void Load(ComboBox tempBox)
            {

                // 休日配列
                string[] sDay = {"01/01元旦", "     成人の日", "02/11建国記念の日", "     春分の日", "04/29昭和の日",
                            "05/03憲法記念日","05/04みどりの日","05/05こどもの日","08/12海の日","     敬老の日",
                            "     秋分の日","     体育の日","11/03文化の日","11/23勤労感謝の日","12/23天皇誕生日",
                            "     振替休日","     国民の休日","     土曜日","     年末年始休暇","     夏季休暇"};

                try
                {
                    comboHoliday cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "Date";

                    foreach (var a in sDay)
                    {
                        cmb1 = new comboHoliday();
                        cmb1.Date = a.Substring(0, 5);
                        int s = a.Length;
                        cmb1.Name = a.Substring(5, s - 5);
                        tempBox.Items.Add(cmb1);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "休日コンボボックスロード");
                }
            }
            
            ///------------------------------------------------------------------------
            /// <summary>
            ///     休日コンボ表示 </summary>
            /// <param name="tempBox">
            ///     コンボボックスオブジェクト</param>
            /// <param name="dt">
            ///     月日</param>
            ///------------------------------------------------------------------------
            public static void selectedIndex(ComboBox tempBox, string dt)
            {
                comboHoliday cmbS = new comboHoliday();
                Boolean Sh = false;

                for (int iX = 0; iX <= tempBox.Items.Count - 1; iX++)
                {
                    tempBox.SelectedIndex = iX;
                    cmbS = (comboHoliday)tempBox.SelectedItem;

                    if (cmbS.Date == dt)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempBox.SelectedIndex = -1;
                }
            }
        }
        
        // 部門コンボボックスクラス
        public class ComboBumon
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            ///----------------------------------------------------------------
            /// <summary>
            ///     ＣＳＶデータから部門コンボボックスにロードする </summary>
            /// <param name="tempObj">
            ///     コンボボックスオブジェクト</param>
            /// <param name="fName">
            ///     ＣＳＶデータファイルパス</param>
            ///----------------------------------------------------------------
            public static void loadBusho(ComboBox tempObj, string dbName)
            {
                try
                {
                    ComboBumon cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "DisplayName";
                    tempObj.ValueMember = "code";

                    // 奉行SQLServer接続文字列取得
                    string sc = sqlControl.obcConnectSting.get(dbName);
                    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                    StringBuilder sb = new StringBuilder();
                    sb.Clear();
                    sb.Append("select DepartmentCode,DepartmentName from tbDepartment ");
                    sb.Append("order by DepartmentCode");

                    SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

                    while (dR.Read())
                    {
                        string dCode = string.Empty;

                        // コンボボックスにセット
                        cmb1 = new ComboBumon();
                        cmb1.ID = string.Empty;

                        if (Utility.NumericCheck(dR["DepartmentCode"].ToString()))
                        {
                            dCode = Utility.StrtoInt(dR["DepartmentCode"].ToString()).ToString().PadLeft(3, '0');
                        }
                        else
                        {
                            dCode = dR["DepartmentCode"].ToString().Trim();
                        }

                        cmb1.DisplayName = dCode + " " + dR["DepartmentName"].ToString();

                        cmb1.Name = dR["DepartmentName"].ToString();
                        cmb1.code = dCode;
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();
                    sdCon.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "部門コンボボックスロード");
                }
            }


            ///----------------------------------------------------------------
            /// <summary>
            ///     ＣＳＶデータから部門コンボボックスに課名をロードする </summary>
            /// <param name="tempObj">
            ///     コンボボックスオブジェクト</param>
            /// <param name="fName">
            ///     ＣＳＶデータファイルパス</param>
            ///----------------------------------------------------------------
            public static void loadKa(ComboBox tempObj, string dbName)
            {
                try
                {
                    ComboBumon cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "DisplayName";
                    tempObj.ValueMember = "code";

                    // 奉行SQLServer接続文字列取得
                    string sc = sqlControl.obcConnectSting.get(dbName);
                    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                    StringBuilder sb = new StringBuilder();
                    sb.Clear();
                    sb.Append("select DepartmentCode,DepartmentName from tbDepartment ");
                    sb.Append("where Right(DepartmentCode, 2) = '00'");
                    sb.Append("order by DepartmentCode");

                    SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

                    while (dR.Read())
                    {
                        string dCode = string.Empty;

                        // コンボボックスにセット
                        cmb1 = new ComboBumon();
                        cmb1.ID = string.Empty;

                        if (Utility.NumericCheck(dR["DepartmentCode"].ToString()))
                        {
                            dCode = Utility.StrtoInt(dR["DepartmentCode"].ToString()).ToString().PadLeft(5, '0');
                        }
                        else
                        {
                            dCode = dR["DepartmentCode"].ToString().Trim();
                        }

                        cmb1.DisplayName = dCode + " " + dR["DepartmentName"].ToString();

                        cmb1.Name = dR["DepartmentName"].ToString();
                        cmb1.code = dCode;
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();
                    sdCon.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "部門コンボボックスロード");
                }
            }
        }

        // 現場コンボボックスクラス
        public class ComboProject
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            ///----------------------------------------------------------------
            /// <summary>
            ///     ＣＳＶデータから部門コンボボックスにロードする </summary>
            /// <param name="tempObj">
            ///     コンボボックスオブジェクト</param>
            /// <param name="fName">
            ///     ＣＳＶデータファイルパス</param>
            ///----------------------------------------------------------------
            public static void load(ComboBox tempObj, string dbName)
            {
                try
                {
                    ComboProject cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "DisplayName";
                    tempObj.ValueMember = "code";

                    // 奉行SQLServer接続文字列取得
                    string sc_ac = sqlControl.obcConnectSting.get(dbName);

                    // 奉行SQLServer接続
                    sqlControl.DataControl sdCon_ac = new sqlControl.DataControl(sc_ac);

                    // プロジェクトデータリーダーを取得する
                    string sqlSTRING = string.Empty;
                    sqlSTRING += "SELECT ProjectCode,ProjectName,ValidDate,InValidDate ";
                    sqlSTRING += "from tbProject ";
                    sqlSTRING += "Order by ProjectCode";

                    SqlDataReader dR = sdCon_ac.free_dsReader(sqlSTRING);

                    while (dR.Read())
                    {
                        string dCode = string.Empty;

                        // コンボボックスにセット
                        cmb1 = new ComboProject();
                        cmb1.ID = string.Empty;
                        
                        dCode = dR["ProjectCode"].ToString().PadLeft(20, '0').Substring(12, 8);

                        cmb1.DisplayName = dCode + " " + dR["ProjectName"].ToString();

                        cmb1.Name = dR["ProjectName"].ToString();
                        cmb1.code = dCode;
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();
                    sdCon_ac.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "現場コンボボックスロード");
                }
            }
        }



        ///------------------------------------------------------------------------
        /// <summary>
        ///     文字列の値が数字かチェックする </summary>
        /// <param name="tempStr">
        ///     検証する文字列</param>
        /// <returns>
        ///     数字:true,数字でない:false</returns>
            ///------------------------------------------------------------------------
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     emptyを"0"に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        ///------------------------------------------------------------------------
        public static string EmptytoZero(string tempStr)
        {
            if (tempStr == string.Empty)
            {
                return "0";
            }
            else
            {
                return tempStr;
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのとき文字型値を返す</returns>
        ///------------------------------------------------------------------------
        public static string NulltoStr(string tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                return tempStr;
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        ///------------------------------------------------------------------------
        public static string NulltoStr(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     文字型をIntへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     Int型の値</returns>
        ///--------------------------------------------------------------------
        public static int StrtoInt(string tempStr)
        {
            if (NumericCheck(tempStr))
            {
                return int.Parse(tempStr);
            }
            else
            {
                return 0;
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     文字型をDoubleへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     double型の値</returns>
        ///--------------------------------------------------------------------
        public static double StrtoDouble(string tempStr)
        {
            if (NumericCheck(tempStr))
            {
                return double.Parse(tempStr);
            }
            else
            {
                return 0;
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     文字型をDecimalへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     Decimal型の値</returns>
        ///--------------------------------------------------------------------
        public static Decimal StrtoDecimal(string tempStr)
        {
            if (NumericCheck(tempStr))
            {
                return Decimal.Parse(tempStr);
            }
            else
            {
                return 0;
            }
        }
        ///-----------------------------------------------------------------------
        /// <summary>
        ///     経過時間を返す </summary>
        /// <param name="s">
        ///     開始時間</param>
        /// <param name="e">
        ///     終了時間</param>
        /// <returns>
        ///     経過時間</returns>
        ///-----------------------------------------------------------------------
        public static TimeSpan GetTimeSpan(DateTime s, DateTime e)
        {
            TimeSpan ts;
            if (s > e)
            {
                TimeSpan j = new TimeSpan(24, 0, 0);
                ts = e + j - s;
            }
            else
            {
                ts = e - s;
            }

            return ts;
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     指定した精度の数値に切り捨てします。</summary>
        /// <param name="dValue">
        ///     丸め対象の倍精度浮動小数点数。</param>
        /// <param name="iDigits">
        ///     戻り値の有効桁数の精度。</param>
        /// <returns>
        ///     iDigits に等しい精度の数値に切り捨てられた数値。</returns>
        /// ------------------------------------------------------------------------
        public static double ToRoundDown(double dValue, int iDigits)
        {
            double dCoef = System.Math.Pow(10, iDigits);

            return dValue > 0 ? System.Math.Floor(dValue * dCoef) / dCoef :
                                System.Math.Ceiling(dValue * dCoef) / dCoef;
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     ファイル選択ダイアログボックスの表示 </summary>
        /// <param name="sTitle">
        ///     タイトル文字列</param>
        /// <param name="sFilter">
        ///     ファイルのフィルター</param>
        /// <returns>
        ///     選択したファイル名</returns>
        ///------------------------------------------------------------------
        public static string userFileSelect(string sTitle, string sFilter)
        {
            DialogResult ret;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //ダイアログボックスの初期設定
            openFileDialog1.Title = sTitle;
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = sFilter;
            //openFileDialog1.Filter = "CSVファイル(*.CSV)|*.csv|全てのファイル(*.*)|*.*";

            //ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return string.Empty;
            }

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return string.Empty;
            }

            return openFileDialog1.FileName;
        }

        public class frmMode
        {
            public int ID { get; set; }

            public int Mode { get; set; }

            public int rowIndex { get; set; }
        }

        public class xlsShain
        {
            public int sCode { get; set; }
            public string sName { get; set; }
            public int bCode { get; set; }
            public string bName { get; set; }
        }
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVファイルを追加モードで出力する</summary>
        /// <param name="sPath">
        ///     出力するパス</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        /// <param name="sFileName">
        ///     CSVファイル名</param>
        ///----------------------------------------------------------------------------
        public static void csvFileWrite(string sPath, string[] arrayData, string sFileName)
        {
            //// ファイル名（タイムスタンプ）
            //string timeStamp = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
            //                     DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
            //                     DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

            //// ファイル名
            //string outFileName = sPath + timeStamp + ".csv";

            //// 出力ファイルが存在するとき
            //if (System.IO.File.Exists(outFileName))
            //{
            //    // 既存のファイルを削除
            //    System.IO.File.Delete(outFileName);
            //}

            // CSVファイル出力
            //System.IO.File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding("shift-jis"));
            System.IO.File.AppendAllLines(sPath, arrayData, Encoding.GetEncoding("shift-Jis"));
        }
        
        ///---------------------------------------------------------------------
        /// <summary>
        ///     任意のディレクトリのファイルを削除する </summary>
        /// <param name="sPath">
        ///     指定するディレクトリ</param>
        /// <param name="sFileType">
        ///     ファイル名及び形式</param>
        /// --------------------------------------------------------------------
        public static void FileDelete(string sPath, string sFileType)
        {
            //sFileTypeワイルドカード"*"は、すべてのファイルを意味する
            foreach (string files in System.IO.Directory.GetFiles(sPath, sFileType))
            {
                // ファイルを削除する
                System.IO.File.Delete(files);
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     文字列を指定文字数をＭＡＸとして返します</summary>
        /// <param name="s">
        ///     文字列</param>
        /// <param name="n">
        ///     文字数</param>
        /// <returns>
        ///     文字数範囲内の文字列</returns>
        /// --------------------------------------------------------------------
        public static string GetStringSubMax(string s, int n)
        {
            string val = string.Empty;

            // 文字間のスペースを除去 2015/03/10
            s = s.Replace(" ", "");

            if (s.Length > n) val = s.Substring(0, n);
            else val = s;

            return val;
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     訂正欄にチェックされているとき「１」されていないとき「０」を返す</summary>
        /// <param name="s">
        ///     文字列</param>
        /// --------------------------------------------------------------------
        public static string getTeiseiCheck(string s)
        {
            string val = string.Empty;

            if (s != string.Empty)
            {
                val = global.FLGON;
            }
            else
            {
                val = global.FLGOFF;
            }

            return val;
        }
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     分記入範囲チェック：0～59の数値及び記入単位 </summary>
        /// <param name="h">
        ///     記入値</param>
        /// <param name="tani">
        ///     記入単位分</param>
        /// <returns>
        ///     正常:true, エラー:false</returns>
        ///------------------------------------------------------------------------------------
        public static bool checkMinSpan(string m, int tani)
        {
            if (!Utility.NumericCheck(m))
            {
                return false;
            }
            else if (int.Parse(m) < 0 || int.Parse(m) > 59)
            {
                return false;
            }
            else if (int.Parse(m) % tani != 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間記入範囲チェック 0～23の数値 </summary>
        /// <param name="h">
        ///     記入値</param>
        /// <param name="maxHour">
        ///     最大記入時間</param>
        /// <returns>
        ///     正常:true, エラー:false</returns>
        ///------------------------------------------------------------------------------------
        public static bool checkHourSpan(string h, int maxHour)
        {
            if (!Utility.NumericCheck(h))
            {
                return false;
            }
            else if (int.Parse(h) < 0 || int.Parse(h) > maxHour)
            {
                return false;
            }
            else
            {
                return true;
            }
        }


        // 勘定奉行データベース接続
        public class SQLDBConnect
        {
            SqlConnection cn = new SqlConnection();

            public SqlConnection Cn
            {
                get
                {
                    return cn;
                }
            }

            /// <summary>
            /// SQLServerへ接続
            /// </summary>
            /// <param name="sConnect">接続文字列</param>
            public SQLDBConnect(string sConnect)
            {
                try
                {
                    // データベース接続文字列
                    cn.ConnectionString = sConnect;
                    cn.Open();
                }

                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     社員情報抽出ＳＱＬ作成 </summary>
        /// <param name="bCode">
        ///     社員コード</param>
        /// <returns>
        ///     ＳＱＬ文字列</returns>
        ///---------------------------------------------------------------
        public static string getEmployee(string bCode)
        {
            string dt = DateTime.Today.ToShortDateString();

            // 社員情報抽出ＳＱＬ
            StringBuilder sb = new StringBuilder();

            //sb.Append("SELECT tbEmployeeBase.EmployeeID, tbHR_DivisionCategory.CategoryCode as zaisekikbn,");
            //sb.Append("tbEmployeeBase.EmployeeNo, tbEmployeeBase.NameKana, tbEmployeeBase.Name,");
            //sb.Append("tbDepartment.DepartmentID, right(replace(tbDepartment.DepartmentCode, ' ', ''), 5) as DepartmentCode, tbDepartment.DepartmentName,");
            //sb.Append("tbEmployeeBase.RetireCorpScheduleDate ");

            //sb.Append("from(((tbEmployeeBase inner join ");

            //sb.Append("(select tbEmployeeMainDutyPersonnelChange.EmployeeID, tbEmployeeMainDutyPersonnelChange.AnnounceDate,");
            //sb.Append("tbEmployeeMainDutyPersonnelChange.BelongID ");

            //sb.Append("from tbEmployeeMainDutyPersonnelChange inner join ");

            //sb.Append("(select EmployeeID, max(AnnounceDate) as AnnounceDate from tbEmployeeMainDutyPersonnelChange ");
            //sb.Append("where AnnounceDate <= '").Append(DateTime.Today.ToShortDateString()).Append("' ");
            //sb.Append("group by EmployeeID) as a ");
            //sb.Append("on(tbEmployeeMainDutyPersonnelChange.EmployeeID = a.EmployeeID) and ");
            //sb.Append("(tbEmployeeMainDutyPersonnelChange.AnnounceDate = a.AnnounceDate) ");
            //sb.Append(") as d ");
            //sb.Append("on tbEmployeeBase.EmployeeID = d.EmployeeID) ");

            //sb.Append("inner join tbDepartment on d.BelongID = tbDepartment.DepartmentID) ");
            //sb.Append("inner join tbHR_DivisionCategory on tbEmployeeBase.BeOnTheRegisterDivisionID = tbHR_DivisionCategory.CategoryID) ");

            //sb.Append("where EmployeeNo = '" + bCode + "' and tbHR_DivisionCategory.CategoryCode <> 2 "); // 2017/05/08 

            //// 在籍区分 <> 2 を外した : 2017/09/28　
            //sb.Append("where EmployeeNo = '" + bCode + "' ");
            //sb.Append("ORDER BY DepartmentCode,tbEmployeeBase.EmployeeNo");
            
            sb.Append("SELECT tbEmployeeBase.EmployeeID, tbHR_DivisionCategory.CategoryCode as zaisekikbn, k.CategoryCode as koyoukbn, ");
            sb.Append("tbEmployeeBase.EmployeeNo, tbEmployeeBase.NameKana, tbEmployeeBase.Name, ");
            sb.Append("tbDepartment.DepartmentID, right(replace(tbDepartment.DepartmentCode, ' ', ''), 3) as DepartmentCode, tbDepartment.DepartmentName, ");
            sb.Append("tbEmployeeBase.RetireCorpScheduleDate, kk.CategoryCode as kyuyoKbn, kk.CategoryName as kyuyoKbnName "); 
            sb.Append("from((((((tbEmployeeBase inner join  ");
            sb.Append("(select tbEmployeeMainDutyPersonnelChange.EmployeeID, tbEmployeeMainDutyPersonnelChange.AnnounceDate, ");
            sb.Append("tbEmployeeMainDutyPersonnelChange.BelongID "); 
            sb.Append("from tbEmployeeMainDutyPersonnelChange inner join "); 

            sb.Append("(select EmployeeID, max(AnnounceDate) as AnnounceDate from tbEmployeeMainDutyPersonnelChange ");
            sb.Append("where AnnounceDate <= '").Append(DateTime.Today.ToShortDateString()).Append("'  ");
            sb.Append("group by EmployeeID) as a  ");
            sb.Append("on(tbEmployeeMainDutyPersonnelChange.EmployeeID = a.EmployeeID) and "); 
            sb.Append("(tbEmployeeMainDutyPersonnelChange.AnnounceDate = a.AnnounceDate)");
            sb.Append(") as d ");

            sb.Append("on tbEmployeeBase.EmployeeID = d.EmployeeID) ");
            sb.Append("inner join tbDepartment on d.BelongID = tbDepartment.DepartmentID) "); 
            sb.Append("inner join tbHR_DivisionCategory on tbEmployeeBase.BeOnTheRegisterDivisionID = tbHR_DivisionCategory.CategoryID) ");
            sb.Append("inner join tbHR_DivisionCategory as k on tbEmployeeBase.EmploymentDivisionID = k.CategoryID) ");            
            sb.Append("inner join ");

            sb.Append("(select tbEmployeeUnitPrice.EmployeeID,tbEmployeeUnitPrice.SalaryDivisionID, max(tbEmployeeUnitPrice.RevisionDate) as rDate  ");
            sb.Append("from tbEmployeeUnitPrice ");
            sb.Append("where RevisionDate <= '").Append(DateTime.Today.ToShortDateString()).Append("' ");
            sb.Append("group by EmployeeID, SalaryDivisionID) as u ");

            sb.Append("on tbEmployeeBase.EmployeeID = u.EmployeeID) ");
            sb.Append("inner join tbHR_DivisionCategory as kk on u.SalaryDivisionID = kk.CategoryID)");

            //sb.Append("where EmployeeNo = '" + bCode + "' and tbHR_DivisionCategory.CategoryCode <> 2 ");
            sb.Append("where EmployeeNo = '" + bCode + "'");

            return sb.ToString();
        }
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     ＣＳＶファイルを出力する</summary>
        /// <param name="sPath">
        ///     出力パス</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        /// <param name="_fileName">
        ///     拡張子を含むファイル名</param>
        /// <param name="overWrite">
        ///     上書きのとき true, 元ファイルをコピーして残す false</param>
        ///----------------------------------------------------------------------------
        public static void txtFileWrite(string sPath, string[] arrayData, string _fileName, bool overWrite)
        {
            // 同名ファイルがあった場合
            if (System.IO.File.Exists(sPath + _fileName))
            {
                if (overWrite)
                {
                    // オーバーライト削除する
                    System.IO.File.Delete(sPath + _fileName);
                }
                else
                {
                    // 付加文字列（タイムスタンプ）
                    string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
                                            DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                            DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

                    // 新ファイル名
                    string outFileName = sPath + newFileName + ".csv";
                    System.IO.File.Move(sPath + _fileName, outFileName);
                }
            }

            // テキストファイル出力
            System.IO.File.WriteAllLines(sPath + _fileName, arrayData, System.Text.Encoding.GetEncoding(932));
        }


        ///--------------------------------------------------------------------------
        /// <summary>
        ///     このコンピュータが出力先マスターに登録されているか調べる</summary>
        /// <returns>
        ///     登録済：true, 未登録：false</returns>
        ///--------------------------------------------------------------------------
        public static string getPcDir()
        {
            string rtn = string.Empty;

            CONYXDataSet dts = new CONYX_OCR.CONYXDataSet();
            CONYXDataSetTableAdapters.出力先ＰＣTableAdapter adp = new CONYXDataSetTableAdapters.出力先ＰＣTableAdapter();

            adp.Fill(dts.出力先ＰＣ);

            try
            {
                if (!dts.出力先ＰＣ.Any(a => a.コンピューター名 == System.Net.Dns.GetHostName()))
                {
                    rtn = string.Empty;
                }
                else
                {
                    var ss = dts.出力先ＰＣ.Single(a => a.コンピューター名 == System.Net.Dns.GetHostName());
                    rtn = ss.登録名;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {

            }

            return rtn;
        }

    }
}
