using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace CONYX_OCR.common
{
    ///------------------------------------------------------------------
    /// <summary>
    ///     給与計算受け渡しデータ作成クラス </summary>     
    ///------------------------------------------------------------------
    class OCROutput
    {
        // 親フォーム
        Form _preForm;

        #region データテーブルインスタンス
        CONYX_CLIDataSet.勤務票ヘッダDataTable _hTbl;
        CONYX_CLIDataSet.勤務票明細DataTable _mTbl;
        CONYX_CLIDataSet.汎用データ作成元DataTable _cTbl;
        int _dtIndi = 0;
        #endregion

        private const string TXTFILENAME = "就業データ";

        CONYX_CLIDataSet _dts = new CONYX_CLIDataSet();

        // 就業奉行汎用データヘッダ項目
        const string H1 = @"""EBAS001""";   // 社員番号
        const string H2 = @"""LTLT001""";   // 日付
        const string H3 = @"""LTLT002""";   // 勤務回
        const string H4 = @"""LTLT003""";   // 勤務体系コード
        const string H5 = @"""LTLT004""";   // 事由コード１
        const string H6 = @"""LTLT005""";   // 事由コード２
        const string H7 = @"""LTLT006""";   // 事由コード３
        const string H8 = @"""LTDT001""";   // 出勤時刻
        const string H9 = @"""LTDT002""";   // 退出時刻
        const string H10 = @"""LTTC001""";   // 休憩時間項目コード
        const string H11 = @"""LTTS001""";   // 休憩時間

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     給与計算用計算用受入データ作成クラスコンストラクタ</summary>
        /// <param name="preFrm">
        ///     親フォーム</param>
        /// <param name="dts">
        ///     CONYX_CLIDataSet</param>
        /// <param name="dtIndi">
        ///     0:24時間制、1:48時間制</param>
        ///--------------------------------------------------------------------------
        public OCROutput(Form preFrm, CONYX_CLIDataSet dts, int dtIndi)
        {
            _preForm = preFrm;
            _dts = dts;
            _cTbl = dts.汎用データ作成元;
            _mTbl = dts.勤務票明細;
            _dtIndi = dtIndi;
        }

        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     就業奉行用受入データ作成</summary>
        ///--------------------------------------------------------------------------------------     
        public void SaveData(string dbName)
        {
            #region 出力配列
            string[] arrayCsv = null;     // 出力配列
            #endregion

            #region 出力件数変数
            int sCnt = 0;   // 社員出力件数
            #endregion

            StringBuilder sb = new StringBuilder();
            Boolean pblFirstGyouFlg = true;
            string wID = string.Empty;
            string hKinmutaikei = string.Empty;

            global gl = new global();

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 出力先フォルダがあるか？なければ作成する
            //string cPath = Properties.Settings.Default.okPath;
            string cPath = global.cnfPath;

            if (!System.IO.Directory.Exists(cPath))
            {
                System.IO.Directory.CreateDirectory(cPath);
            }

            try
            {
                //オーナーフォームを無効にする
                _preForm.Enabled = false;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = _preForm;
                frmP.Show();

                int rCnt = 1;
                int saSum = 0;
                DateTime saDay = global.NODATE;
                int kaisu = 1;
                DateTime chkDt = DateTime.Today;

                // 伝票最初行フラグ
                pblFirstGyouFlg = true;

                // 勤務票データ取得
                CONYX_CLIDataSetTableAdapters.汎用データ作成元TableAdapter adp = new CONYX_CLIDataSetTableAdapters.汎用データ作成元TableAdapter();
                adp.Fill(_cTbl);

                foreach (var r in _cTbl)
                {
                    // 存在しない日付は対象外
                    if (r.歴日付 == global.NODATE)
                    {
                        continue;
                    }

                    // 勤務実績のないデータは対象外
                    if (r.勤務体系コード == string.Empty && r.開始時 == string.Empty && 
                        r.開始分 == string.Empty && r.退勤時 == string.Empty && r.退勤分 == string.Empty && 
                        r.休憩時 == string.Empty && r.休憩分 == string.Empty && r.実労時 == string.Empty && 
                        r.実労分 == string.Empty && r.事由１ == string.Empty && r.事由2 == string.Empty)
                    {
                        continue;
                    } 
                    
                    // プログレスバー表示
                    frmP.Text = "就業奉行用受入データ作成中です・・・" + rCnt.ToString() + "/" + _cTbl.Count().ToString();
                    frmP.progressValue = rCnt * 100 / _cTbl.Count();
                    frmP.ProgressStep();

                    // ヘッダファイル出力
                    if (pblFirstGyouFlg == true)
                    {
                        sb.Clear();
                        sb.Append(H1).Append(",");      // 社員番号
                        sb.Append(H2).Append(",");      // 日付
                        sb.Append(H3).Append(",");      // 勤務回
                        sb.Append(H4).Append(",");      // 勤務体系コード
                        sb.Append(H5).Append(",");      // 事由１
                        sb.Append(H6).Append(",");      // 事由２
                        //sb.Append(H7).Append(",");      // 事由３
                        sb.Append(H8).Append(",");      // 出勤時刻
                        sb.Append(H9).Append(",");      // 退勤時刻
                        sb.Append(H10).Append(",");     // 休憩時間項目コード
                        sb.Append(H11);                 // 休憩時間

                        // 配列にデータを出力
                        sCnt++;
                        Array.Resize(ref arrayCsv, sCnt);
                        arrayCsv[sCnt - 1] = sb.ToString();
                    }

                    sb.Clear();

                    // 社員番号
                    sb.Append(r.社員番号).Append(",");

                    // 日付
                    sb.Append(r.歴日付.ToShortDateString()).Append(",");

                    // 勤務回
                    if (saSum == r.社員番号 && saDay == r.歴日付)
                    {
                        kaisu++;
                    }
                    else
                    {
                        kaisu = 1;
                    }

                    sb.Append(kaisu).Append(",");

                    // 勤務体系（シフト）コード
                    hKinmutaikei = r.勤務体系コード.ToString();
                    sb.Append(hKinmutaikei).Append(",");

                    // 事由
                    sb.Append(r.事由１).Append(",");
                    sb.Append(r.事由2).Append(",");

                    // 開始時刻
                    if (r.開始時 == string.Empty && r.開始分 == string.Empty)
                    {
                        sb.Append(",");
                    }
                    else
                    {
                        sb.Append(r.開始時 + ":" + r.開始分.PadLeft(2, '0')).Append(",");
                    }

                    //退出時刻
                    int taikinH = Utility.StrtoInt(r.退勤時);

                    // 就業奉行が24時間制のとき
                    if (_dtIndi == 0)
                    {
                        if (taikinH >= 24)
                        {
                            // 先頭に「翌日」を付加
                            sb.Append("翌日");
                            taikinH -= 24;
                        }
                    }

                    if (r.退勤時 == string.Empty && r.退勤分 == string.Empty)
                    {
                        sb.Append(",");
                    }
                    else
                    {
                        sb.Append(taikinH + ":" + r.退勤分.PadLeft(2, '0')).Append(",");
                    }

                    // 休憩時間
                    sb.Append("012,");

                    if (r.休憩時 == string.Empty && r.休憩分 == string.Empty)
                    {
                        //sb.Append(",");
                        sb.Append("0:00");
                    }
                    else
                    {
                        //sb.Append(r.休憩時 + ":" + r.休憩分.PadLeft(2, '0')).Append(",");
                        sb.Append(r.休憩時 + ":" + r.休憩分.PadLeft(2, '0'));
                    }

                    string uCSV = sb.ToString();

                    
                    // 配列にデータを格納します
                    sCnt++;
                    Array.Resize(ref arrayCsv, sCnt);
                    arrayCsv[sCnt - 1] = uCSV;

                    // データ件数加算
                    rCnt++;

                    pblFirstGyouFlg = false;

                    saSum = r.社員番号;
                    saDay = r.歴日付;
                }

                // 勤怠CSVファイル出力
                if (arrayCsv != null)
                {
                    txtFileWrite(cPath, arrayCsv);
                }

                // いったんオーナーをアクティブにする
                _preForm.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                _preForm.Enabled = true;
            }
            catch (Exception e)
            {
                MessageBox.Show("就業奉行受入データ作成中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                //if (OutData.sCom.Connection.State == ConnectionState.Open) OutData.sCom.Connection.Close();

                if (sdCon.Cn.State == System.Data.ConnectionState.Open)
                {
                    sdCon.Close();
                }
            }
        }
        
        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象シフトコードの開始時刻と終了時刻を取得する 
        ///     : 日替わり時刻を取得 2018/02/06 </summary>
        /// <param name="_dbName">
        ///     データベース名</param>
        /// <param name="sftCode">
        ///     シフトコード </param>
        /// <param name="sTime">
        ///     開始時刻</param>
        /// <param name="eTime">
        ///     終了時刻</param>
        /// <param name="changeTime">
        ///     日替わり時刻</param>
        ///----------------------------------------------------------------------------------
        private void GetSftTime(string sftCode, out string sTime, out string eTime, out DateTime changeTime, sqlControl.DataControl sdCon)
        {
            // 対象のシフトコード取得する
            DateTime sDt = DateTime.Now;
            DateTime eDt = DateTime.Now;
            DateTime cDt = DateTime.Now;    // 2018/02/06

            // 勤務体系（シフト）コード情報取得
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("SELECT tbLaborSystem.LaborSystemID,tbLaborSystem.LaborSystemCode,");
            sb.Append("tbLaborSystem.LaborSystemName,tbLaborTimeSpanRule.StartTime,");
            sb.Append("tbLaborTimeSpanRule.EndTime,tbLaborSystem.DayChangeTime ");
            sb.Append("FROM tbLaborSystem inner join tbLaborTimeSpanRule ");
            sb.Append("on tbLaborSystem.LaborSystemID = tbLaborTimeSpanRule.LaborSystemID ");
            sb.Append("where tbLaborTimeSpanRule.LaborTimeSpanRuleType = 1 ");
            sb.Append("and tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                sDt = DateTime.Parse(dR["StartTime"].ToString());
                eDt = DateTime.Parse(dR["EndTime"].ToString());
                cDt = DateTime.Parse(dR["DayChangeTime"].ToString());   // 2018/02/06
                break;
            }

            dR.Close();

            // 開始時刻
            sTime = sDt.Hour.ToString() + ":" + sDt.Minute.ToString().PadLeft(2, '0');

            // 終了時刻
            eTime = string.Empty;
            if (sDt.Day < eDt.Day)
            {
                // 翌日のとき
                eTime = "翌日";
            }

            eTime += eDt.Hour.ToString() + ":" + eDt.Minute.ToString().PadLeft(2, '0');

            // 2018/02/06
            changeTime = cDt;
        }
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     配列にテキストデータをセットする </summary>
        /// <param name="array">
        ///     社員、パート、出向社員の各配列</param>
        /// <param name="cnt">
        ///     拡張する配列サイズ</param>
        /// <param name="txtData">
        ///     セットする文字列</param>
        ///----------------------------------------------------------------------------
        private void txtArraySet(string [] array, int cnt, string txtData)
        {
            Array.Resize(ref array, cnt);   // 配列のサイズ拡張
            array[cnt - 1] = txtData;       // 文字列のセット
        }
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     テキストファイルを出力する</summary>
        /// <param name="outFilePath">
        ///     出力するフォルダ</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        ///----------------------------------------------------------------------------
        private void txtFileWrite(string sPath, string [] arrayData)
        {
            // 付加文字列（タイムスタンプ）
            string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
                                    DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                    DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0'); 

            // ファイル名
            string outFileName = sPath + TXTFILENAME + newFileName + ".csv";
            
            // テキストファイル出力
            File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding(932));
        }
    }
}
