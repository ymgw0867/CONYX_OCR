using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;

namespace CONYX_OCR.common
{
    class OCRData
    {
        public OCRData(string dbName)
        {
            _dbName = dbName;           // 人事給与
        }

        // 奉行シリーズデータ領域データベース名
        string _dbName = string.Empty;
        string _dbName_ac = string.Empty;

        //common.xlsData bs;

        #region エラー項目番号プロパティ
        //---------------------------------------------------
        //          エラー情報
        //---------------------------------------------------

        enum errCode
        {
            eNothing, eYearMonth, eMonth, eDay, eKinmuTaikeiCode
        }

        /// <summary>
        ///     エラーヘッダ行RowIndex</summary>
        public int _errHeaderIndex { get; set; }

        /// <summary>
        ///     エラー項目番号</summary>
        public int _errNumber { get; set; }

        /// <summary>
        ///     エラー明細行RowIndex </summary>
        public int _errRow { get; set; }

        /// <summary> 
        ///     エラーメッセージ </summary>
        public string _errMsg { get; set; }

        /// <summary> 
        ///     エラーなし </summary>
        public int eNothing = 0;

        /// <summary>
        ///     エラー項目 = 確認チェック </summary>
        public int eDataCheck = 35;

        /// <summary> 
        ///     エラー項目 = 対象年月日 </summary>
        public int eYearMonth = 1;

        /// <summary> 
        ///     エラー項目 = 対象月 </summary>
        public int eMonth = 2;

        /// <summary> 
        ///     エラー項目 = 日 </summary>
        public int eDay = 3;

        /// <summary> 
        ///     エラー項目 = 出勤状況 </summary>
        public int eShukkinStatus = 4;

        /// <summary> 
        ///     エラー項目 = 個人番号 </summary>
        public int eShainNo = 5;
        public int eShainNo2 = 27;

        /// <summary> 
        ///     エラー項目 = 勤務体系 </summary>
        public int eKinmuTaikeiCode = 6;

        /// <summary> 
        ///     エラー項目 = 単価振替区分 </summary>
        public int eTankaKbn = 31;

        /// <summary> 
        ///     エラー項目 = 事由１ </summary>
        public int eJiyu1 = 7;

        /// <summary> 
        ///     エラー項目 = 事由２ </summary>
        public int eJiyu2 = 8;

        /// <summary> 
        ///     エラー項目 = 残業申請書 </summary>
        public int eCFlg = 9;

        /// <summary>
        ///     エラー項目 = 交通区分 </summary>
        public int eKotsuKbn = 10;

        /// <summary>
        ///     エラー項目 = 走行距離 </summary>
        public int eSoukou = 11;

        /// <summary> 
        ///     エラー項目 = 同乗人数 </summary>
        public int eDoujyoNin = 12;
        
        /// <summary> 
        ///     エラー項目 = 開始時 </summary>
        public int eSH = 13;

        /// <summary> 
        ///     エラー項目 = 開始分 </summary>
        public int eSM = 14;

        /// <summary> 
        ///     エラー項目 = 終了時 </summary>
        public int eEH = 15;

        /// <summary> 
        ///     エラー項目 = 終了分 </summary>
        public int eEM = 16;

        /// <summary> 
        ///     エラー項目 = 休憩 </summary>
        //public int eRest = 17;

        /// <summary> 
        ///     エラー項目 = 実働時間 </summary>
        public int eWh = 18;

        /// <summary> 
        ///     エラー項目 = 実働分 </summary>
        public int eWm = 19;

        /// <summary> 
        ///     エラー項目 = 勤務時間区分 </summary>
        public int eKinmuKbn = 20;

        /// <summary> 
        ///     エラー項目 = 交通費 </summary>
        public int eKotsuhi = 21;

        /// <summary> 
        ///     エラー項目 = 公休日数 </summary>
        public int eKoukyuDays = 22;

        /// <summary> 
        ///     エラー項目 = 休憩時間・分 </summary>
        public int eRh = 23;
        public int eRm = 28;

        /// <summary> 
        ///     エラー項目 = 清掃出勤簿承認印 </summary>
        public int eShouninIn = 24;

        /// <summary> 
        ///     エラー項目 = 警備報告書確認印 </summary>
        public int eKakuninIn = 25;

        /// <summary> 
        ///     エラー項目 = 応援分 </summary>
        public int eOuenM = 26;

        /// <summary> 
        ///     エラー項目 = 応援分 </summary>
        public int eOuenIP = 32;
        public int eOuenIP2 = 33;

        /// <summary> 
        ///     エラー項目 = 応援移動票と勤怠データＩ／Ｐ票 </summary>
        public int eIpOuen = 34;

        #endregion
        
        #region 警告項目
        ///     <!--警告項目配列 -->
        public int[] warArray = new int[6];

        /// <summary>
        ///     警告項目番号</summary>
        public int _warNumber { get; set; }

        /// <summary>
        ///     警告明細行RowIndex </summary>
        public int _warRow { get; set; }

        /// <summary> 
        ///     警告項目 = 勤怠記号1&2 </summary>
        public int wKintaiKigou = 0;

        /// <summary> 
        ///     警告項目 = 開始終了時分 </summary>
        public int wSEHM = 1;

        /// <summary> 
        ///     警告項目 = 時間外時分 </summary>
        public int wZHM = 2;

        /// <summary> 
        ///     警告項目 = 深夜勤務時分 </summary>
        public int wSIHM = 3;

        /// <summary> 
        ///     警告項目 = 休日出勤時分 </summary>
        public int wKSHM = 4;

        /// <summary> 
        ///     警告項目 = 出勤形態 </summary>
        public int wShukeitai = 5;

        #endregion

        #region フィールド定義
        /// <summary> 
        ///     警告項目 = 時間外1.25時 </summary>
        public int [] wZ125HM = new int[global.MAX_GYO];

        /// <summary> 
        ///     実働時間 </summary>
        public double _workTime;

        /// <summary> 
        ///     深夜稼働時間 </summary>
        public double _workShinyaTime;
        #endregion
        

        #region 時間チェック記号定数
        private const string cHOUR = "H";           // 時間をチェック
        private const string cMINUTE = "M";         // 分をチェック
        private const string cTIME = "HM";          // 時間・分をチェック
        #endregion

        private const string WKSPAN0750 = "7時間50分";
        private const string WKSPAN0755 = "7時間55分";
        private const string WKSPAN0800 = "8時間";
        private const string WKSPAN_KYUJITSU = "休日出勤";

        // 休憩時間
        private const Int64 RESTTIME0750 = 60;      // 7時間50分
        private const Int64 RESTTIME0755 = 65;      // 7時間55分
        private const Int64 RESTTIME0800 = 60;      // 8時間

        // テーブルアダプターマネージャーインスタンス
        CONYX_CLIDataSetTableAdapters.TableAdapterManager adpMn = new CONYX_CLIDataSetTableAdapters.TableAdapterManager();
        //CBSDataSetTableAdapters.休日TableAdapter kAdp = new CBSDataSetTableAdapters.休日TableAdapter();
        
        ///-----------------------------------------------------------------------
        /// <summary>
        ///     CSVデータをMDBに登録する：DataSet Version </summary>
        /// <param name="_InPath">
        ///     CSVデータパス</param>
        /// <param name="frmP">
        ///     プログレスバーフォームオブジェクト</param>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="dbName">
        ///     データ領域データベース名</param>
        ///-----------------------------------------------------------------------
        public void CsvToMdb(string _inPath, frmPrg frmP, string dbName)
        {
            string headerKey = string.Empty;    // ヘッダキー
            int shopCode = 0;     // 店舗コード

            // テーブルセットオブジェクト
            CONYX_CLIDataSet dts = new CONYX_CLIDataSet();

            try
            {
                // 勤務表ヘッダデータセット読み込み
                CONYX_CLIDataSetTableAdapters.勤務票ヘッダTableAdapter hAdp = new CONYX_CLIDataSetTableAdapters.勤務票ヘッダTableAdapter();
                adpMn.勤務票ヘッダTableAdapter = hAdp;
                adpMn.勤務票ヘッダTableAdapter.Fill(dts.勤務票ヘッダ);

                // 勤務表明細データセット読み込み
                CONYX_CLIDataSetTableAdapters.勤務票明細TableAdapter iAdp = new CONYX_CLIDataSetTableAdapters.勤務票明細TableAdapter();
                adpMn.勤務票明細TableAdapter = iAdp;
                adpMn.勤務票明細TableAdapter.Fill(dts.勤務票明細);

                // 対象CSVファイル数を取得
                string [] t = System.IO.Directory.GetFiles(_inPath, "*.csv");
                int cLen = t.Length;

                //CSVデータをMDBへ取込
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_inPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + cLen.ToString();
                    frmP.progressValue = cCnt * 100 / cLen;
                    frmP.ProgressStep();

                    ////////OCR処理対象のCSVファイルかファイル名の文字数を検証する
                    //////string fn = Path.GetFileName(files);

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // ヘッダ行
                        if (stCSV[0] == "*")
                        {
                            // ヘッダーキー取得
                            headerKey = Utility.GetStringSubMax(stCSV[1].Trim(), 17);
                            
                            // データセットに勤務票ヘッダデータを追加する
                            dts.勤務票ヘッダ.Add勤務票ヘッダRow(setNewHeadRecRow(dts, stCSV));
                        }
                        else　// 明細行
                        {
                            // データセットに勤務表明細データを追加する
                            dts.勤務票明細.Add勤務票明細Row(setNewItemRecRow(dts, headerKey, stCSV, shopCode));
                        }
                    }
                }

                // ローカルのデータベースを更新
                adpMn.UpdateAll(dts);

                //CSVファイルを削除する
                foreach (string files in System.IO.Directory.GetFiles(_inPath, "*.csv"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票CSVインポート処理", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用勤務票ヘッダRowオブジェクトを作成する </summary>
        /// <param name="tblSt">
        ///     テーブルセット</param>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <returns>
        ///     追加する勤務票ヘッダRowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private CONYX_CLIDataSet.勤務票ヘッダRow setNewHeadRecRow(CONYX_CLIDataSet tblSt, string[] stCSV)
        {
            CONYX_CLIDataSet.勤務票ヘッダRow r = tblSt.勤務票ヘッダ.New勤務票ヘッダRow();
            r.ID = Utility.GetStringSubMax(stCSV[1].Trim(), 17);
            r.年 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[2].Trim().Replace("-", ""), 4));
            r.月 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[3].Trim().Replace("-", ""), 2));
            r.社員番号 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[4].Trim().Replace("-", ""), 8));
            r.社員名 = string.Empty;
            r.枚数 = string.Empty;
            r.画像名 = Utility.GetStringSubMax(stCSV[1].Trim(), 21);
            r.確認 = global.flgOff;
            r.備考 = string.Empty;
            r.会社領域名 = _dbName;
            r.編集アカウント = global.flgOff;
            r.更新年月日 = DateTime.Now;

            return r;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用勤務票明細Rowオブジェクトを作成する </summary>
        /// <param name="tblSt">
        ///     テーブルセットオブジェクト</param>
        /// <param name="headerKey">
        ///     ヘッダキー</param>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <param name="sShopCode">
        ///     店舗コード</param>
        /// <returns>
        ///     追加する勤務票明細Rowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private CONYX_CLIDataSet.勤務票明細Row setNewItemRecRow(CONYX_CLIDataSet tblSt, string headerKey, string[] stCSV, int sShopCode)
        {
            CONYX_CLIDataSet.勤務票明細Row r = tblSt.勤務票明細.New勤務票明細Row();

            r.ヘッダID = headerKey;
            r.日 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[0].Trim().Replace("-", ""), 2));
            r.勤務体系コード = Utility.GetStringSubMax(stCSV[1].Trim().Replace("-", ""), 4);
            r.残業休日出勤申請 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[2].Trim().Replace("-", ""), 1));
            r.開始時 = Utility.GetStringSubMax(stCSV[3].Trim().Replace("-", ""), 2);
            r.開始分 = Utility.GetStringSubMax(stCSV[4].Trim().Replace("-", ""), 2);
            r.退勤時 = Utility.GetStringSubMax(stCSV[5].Trim().Replace("-", ""), 2);
            r.退勤分 = Utility.GetStringSubMax(stCSV[6].Trim().Replace("-", ""), 2);
            r.休憩時 = Utility.GetStringSubMax(stCSV[7].Trim().Replace("-", ""), 1);
            r.休憩分 = Utility.GetStringSubMax(stCSV[8].Trim().Replace("-", ""), 2);
            r.実労時 = Utility.GetStringSubMax(stCSV[9].Trim().Replace("-", ""), 2);
            r.実労分 = Utility.GetStringSubMax(stCSV[10].Trim().Replace("-", ""), 2);
            r.事由１ = Utility.GetStringSubMax(stCSV[11].Trim().Replace("-", ""), 2);
            r.事由2 = Utility.GetStringSubMax(stCSV[12].Trim().Replace("-", ""), 2);
            r.事由2 = Utility.GetStringSubMax(stCSV[12].Trim().Replace("-", ""), 2);
            r.編集アカウント = global.flgOff;
            r.歴日付 = global.NODATE;
            
            return r;
        }

        ///----------------------------------------------------------------------------------------
        /// <summary>
        ///     値1がemptyで値2がNot string.Empty のとき "0"を返す。そうではないとき値1をそのまま返す</summary>
        /// <param name="str1">
        ///     値1：文字列</param>
        /// <param name="str2">
        ///     値2：文字列</param>
        /// <returns>
        ///     文字列</returns>
        ///----------------------------------------------------------------------------------------
        private string hmStrToZero(string str1, string str2)
        {
            string rVal = str1;
            if (str1 == string.Empty && str2 != string.Empty)
                rVal = "0";

            return rVal;
        }


        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     勤怠データエラーチェックメイン処理。
        ///     エラーのときOCRDataクラスのヘッダ行インデックス、フィールド番号、明細行インデックス、
        ///     エラーメッセージが記録される </summary>
        /// <param name="sIx">
        ///     開始ヘッダ行インデックス</param>
        /// <param name="eIx">
        ///     終了ヘッダ行インデックス</param>
        /// <param name="frm">
        ///     親フォーム</param>
        /// <param name="dts">
        ///     データセット</param>
        /// <returns>
        ///     True:エラーなし、false:エラーあり</returns>
        ///-----------------------------------------------------------------------------------------------
        public Boolean errCheckMain(int sIx, int eIx, Form frm, CONYX_CLIDataSet dts, string[] cID)
        {
            int rCnt = 0;

            // オーナーフォームを無効にする
            frm.Enabled = false;

            // プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = frm;
            frmP.Show();

            // レコード件数取得
            int cTotal = dts.勤務票ヘッダ.Rows.Count;

            // 出勤簿データ読み出し
            Boolean eCheck = true;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);

            // 奉行SQLServer接続
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            try
            {
                for (int i = 0; i < cTotal; i++)
                {
                    //データ件数加算
                    rCnt++;

                    //プログレスバー表示
                    frmP.Text = "エラーチェック実行中　" + rCnt.ToString() + "/" + cTotal.ToString();
                    frmP.progressValue = rCnt * 100 / cTotal;
                    frmP.ProgressStep();

                    //指定範囲ならエラーチェックを実施する：（i:行index）
                    if (i >= sIx && i <= eIx)
                    {
                        // 勤務票ヘッダ行のコレクションを取得します
                        CONYX_CLIDataSet.勤務票ヘッダRow r = dts.勤務票ヘッダ.Single(a => a.ID == cID[i]);

                        // エラーチェック実施
                        eCheck = errCheckData(dts, r, sdCon);

                        if (!eCheck)　//エラーがあったとき
                        {
                            _errHeaderIndex = i;     // エラーとなったヘッダRowIndex
                            break;
                        }
                    }
                }

                return eCheck;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return eCheck;
            }
            finally
            {
                // いったんオーナーをアクティブにする
                frm.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                frm.Enabled = true;

                // 奉行SQLServer接続コネクション閉じる
                sdCon.Close();
            }
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     エラー情報を取得します </summary>
        /// <param name="eID">
        ///     エラーデータのID</param>
        /// <param name="eNo">
        ///     エラー項目番号</param>
        /// <param name="eRow">
        ///     エラー明細行</param>
        /// <param name="eMsg">
        ///     表示メッセージ</param>
        ///---------------------------------------------------------------------------------
        private void setErrStatus(int eNo, int eRow, string eMsg)
        {
            //errHeaderIndex = eHRow;
            _errNumber = eNo;
            _errRow = eRow;
            _errMsg = eMsg;
        }


        ///-----------------------------------------------------------------------------------------------
        /// <summary>
        ///     項目別エラーチェック。
        ///     エラーのときヘッダ行インデックス、フィールド番号、明細行インデックス、エラーメッセージが記録される </summary>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="r">
        ///     勤務票ヘッダ行コレクション</param>
        /// <returns>
        ///     エラーなし：true, エラー有り：false</returns>
        ///-----------------------------------------------------------------------------------------------
        public Boolean errCheckData(CONYX_CLIDataSet dts, CONYX_CLIDataSet.勤務票ヘッダRow r, sqlControl.DataControl sdCon)
        {
            // 確認チェック
            if (r.確認 == global.flgOff)
            {
                setErrStatus(eDataCheck, 0, "未確認の出勤簿です");
                return false;
            }

            // 対象年月
            if (r.年 != global.cnfYear)
            {
                setErrStatus(eYearMonth, 0, "設定された処理年月（" + global.cnfYear + "年" + global.cnfMonth + "月）と異なっています");
                return false;
            }

            if (r.月 != global.cnfMonth)
            {
                setErrStatus(eMonth, 0, "設定された処理年月（" + global.cnfYear + "年" + global.cnfMonth + "月）と異なっています");
                return false;
            }
            
            // 社員番号：数字以外のとき
            if (!Utility.NumericCheck(Utility.NulltoStr(r.社員番号)))
            {
                setErrStatus(eShainNo, 0, "社員番号が入力されていません");
                return false;
            }

            // 登録済み社員番号マスター検証
            if (!chkShainCode(r.社員番号.ToString(), sdCon))
            {
                setErrStatus(eShainNo, 0, "マスター未登録の社員番号です");
                return false;
            }

            //// 同じスタッフ番号の出勤簿が複数存在するときエラー
            //if (!getSameNumber(dts, r.社員番号))
            //{
            //    setErrStatus(eShainNo, 0, "同じスタッフ番号の出勤簿が複数あります");
            //    return false;
            //}

            // 勤務実績チェック
            if (!errCheckNoWork(r)) return false;

            int iX = 0;

            // 勤務票明細データ行を取得
            List<CONYX_CLIDataSet.勤務票明細Row> mList = dts.勤務票明細.Where(a => a.ヘッダID == r.ID).OrderBy(a => a.ID).ToList();

            foreach (var m in mList)
            {
                // 行数
                iX++;

                // 無記入の行はチェック対象外とする
                if (m.開始時 == string.Empty && m.開始分 == string.Empty && 
                    m.退勤時 == string.Empty && m.退勤分 == string.Empty && 
                    m.休憩時 == string.Empty && m.休憩分 == string.Empty &&
                    m.実労時 == string.Empty && m.実労分 == string.Empty &&
                    m.勤務体系コード == string.Empty && m.残業休日出勤申請 == global.flgOff &&
                    m.事由１ == string.Empty && m.事由2 == string.Empty)
                {
                    continue;
                }

                // 日付
                if (!errCheckDay(r, m, "日付", iX))
                {
                    return false;
                }

                // 勤務体系コードマスター登録済み検証
                if (!chkKinmuTaikeiCode(m, "勤務体系コード", iX, sdCon))
                {
                    return false;
                }

                string[] mJiyu = { m.事由１, m.事由2 };

                // 事由区分配列取得：2018/08/10
                OCR.clsAlldayAnotherData ji = new OCR.clsAlldayAnotherData(mJiyu);
                string [] rd = ji.getLaborReasonDivision(m, sdCon);

                // 開始時刻・終業時刻チェック
                if (!errCheckTime(m, "出退時間", global.cnfMarume, iX, rd))
                {
                    return false;
                }

                // 休憩時間
                if (!errCheckRestTime(m, "休憩時間", iX, global.cnfMarume))
                {
                    return false;
                }

                // 実働時間
                if (!errCheckWorkTime(m, "実働時間", iX))
                {
                    return false;
                }
                
                // 残業休日出勤申請書チェック
                // 普通残業
                if (isZangyo(sdCon, m) && m.残業休日出勤申請 == global.flgOff)
                {
                    setErrStatus(eCFlg, iX - 1, "休出、または普通残業に該当しますが休日出勤申請が未チェックです");
                    return false;
                }

                // 早出残業
                if (isHayade(sdCon, m) && m.残業休日出勤申請 == global.flgOff)
                {
                    setErrStatus(eCFlg, iX - 1, "早出に該当しますが休日出勤申請が未チェックです");
                    return false;
                }

                // 事由チェック
                string[] eJiyu = { eJiyu1.ToString(), eJiyu2.ToString() };
                int errNum = 0;

                // 事由コードマスター登録済み検証
                OCR.clsJiyuHas jiyu = new OCR.clsJiyuHas(mJiyu);

                if (!jiyu.isHasRows(out errNum, sdCon))
                {
                    setErrStatus(Utility.StrtoInt(eJiyu[errNum]), iX - 1, "マスター未登録の事由です");
                    return false;
                }

                // 取得単位「終日」事由と他の取得単位あり事由の併記はエラー
                OCR.clsJiyuAllDay jiyu2 = new OCR.clsJiyuAllDay(mJiyu);
                if (!jiyu2.isAllDayAnotherDay(out errNum, sdCon))
                {
                    setErrStatus(Utility.StrtoInt(eJiyu[errNum]), iX - 1, "取得単位が「終日」の事由と他の事由は同時に記入出来ません");
                    return false;
                }

                // 取得単位「終日」事由と他勤務項目記入エラー
                OCR.clsAlldayAnotherData jiyu3 = new OCR.clsAlldayAnotherData(mJiyu);
                if (!jiyu3.isAlldayAnotherData(m, sdCon, out errNum))
                {
                    int eNum = 0;

                    if (errNum == 1)
                    {
                        eNum = eSH;
                    }

                    if (errNum == 2)
                    {
                        eNum = eSM;
                    }

                    if (errNum == 3)
                    {
                        eNum = eEH;
                    }

                    if (errNum == 4)
                    {
                        eNum = eEM;
                    }

                    if (errNum == 5)
                    {
                        eNum = eRh;
                    }

                    if (errNum == 6)
                    {
                        eNum = eRm;
                    }

                    if (errNum == 7)
                    {
                        eNum = eWh;
                    }

                    if (errNum == 8)
                    {
                        eNum = eWm;
                    }
                    
                    setErrStatus(eNum, iX - 1, "取得単位が「終日」の事由で他の項目が記入されています");
                    return false;
                }

                // 「前半」「後半」の重複記入エラー
                OCR.clsJiyuDiv jiyu4 = new OCR.clsJiyuDiv(mJiyu);
                if (!jiyu4.isJiyuDiv(sdCon, out errNum))
                {
                    setErrStatus(Utility.StrtoInt(eJiyu[0]), iX - 1, "「前半」「後半」の事由が重複記入されています");
                    return false;
                }

                // 事由取得単位配列取得：2018/08/10
                string[] st = ji.getAcquireUnit(m, sdCon);

                // 明細記入チェック
                if (!errCheckRow(m, "出勤簿内容", iX, st))
                {
                    return false;
                }

                //// 終日以外の事由記入で勤務体系が未記入のときエラー
                //OCR.clsHanDivShift jiyu5 = new OCR.clsHanDivShift(mJiyu);
                //if (!jiyu5.isHanDivShift(m, sdCon))
                //{
                //    setErrStatus(eKinmuTaikeiCode, iX - 1, "勤務体系コードが未入力です");
                //    return false;
                //}
            }

            return true;
        }
        

        ///-------------------------------------------------------------
        /// <summary>
        ///     基本就業時間帯が時間表記されているか </summary>
        /// <param name="sHH">
        ///     時</param>
        /// <param name="sMM">
        ///     分</param>
        /// <returns>
        ///     true:時間表記, false:時間表記でない</returns>
        ///-------------------------------------------------------------
        private bool isKihonTime(string sHH, string sMM)
        {
            DateTime sHHMM;
            bool rtn = false;

            if (DateTime.TryParse(sHH + ":" + sMM, out sHHMM))
            {
                rtn = true;
            }

            return rtn;
        }



        //private bool chkJiyuHas(CBSDataSet.勤務票明細Row m, string mJiyu)
        //{
        //    OCR.clsJiyuHas jiyu = new OCR.clsJiyuHas(mJiyu, _dbName);

        //    // マスター登録チェック
        //    if (!jiyu.isHasRows())
        //    {
        //        return false;
        //    }
        //    else
        //    {
        //        return true;
        //    }
        //}

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間記入チェック </summary>
        /// <param name="m">
        ///     CONYX_CLIDataSet.勤務票明細Row</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="Tani">
        ///     分記入単位</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="rd">
        ///     事由区分配列</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckTime(CONYX_CLIDataSet.勤務票明細Row m, string tittle, int Tani, int iX, string [] rd)
        {
            // 出勤時間と退勤時間
            string sTimeW = m.開始時.Trim() + m.開始分.Trim();
            string eTimeW = m.退勤時.Trim() + m.退勤分.Trim();

            // 出勤時刻記入あり、退勤時刻記入なし
            if (sTimeW != string.Empty && eTimeW == string.Empty)
            {
                setErrStatus(eEH, iX - 1, "退勤時刻が未入力です");
                return false;
            }

            // 直行直帰、出張を条件に入れないのでとりあえずコメント化 : 2018/08/13
            //// 直帰、直行直帰、出張以外で退勤時刻が未入力のとき : 2018/08/10
            //bool ch = false;
            //for (int i = 0; i < rd.Length; i++)
            //{
            //    if (Utility.StrtoInt(rd[i]) == global.JIYUKBN_CHOKKI || 
            //        Utility.StrtoInt(rd[i]) == global.JIYUKBN_CHOKKOUCHOKKI || 
            //        Utility.StrtoInt(rd[i]) == global.JIYUKBN_SHUCCHOU)
            //    {
            //        ch = true;
            //        break;
            //    }
            //}

            //if (!ch)
            //{
            //    // 出勤時刻記入あり、退勤時刻記入なし
            //    if (sTimeW != string.Empty && eTimeW == string.Empty)
            //    {
            //        setErrStatus(eEH, iX - 1, "退勤時刻が未入力です");
            //        return false;
            //    }
            //}


            // 出勤時刻記入なし、退勤時刻記入あり
            if (sTimeW == string.Empty && eTimeW != string.Empty)
            {
                setErrStatus(eSH, iX - 1, "出勤時刻が未入力です");
                return false;
            }
            
            // 直行直帰、出張を条件に入れないのでとりあえずコメント化 : 2018/08/13
            //// 直行、直行直帰、出張以外で出勤時刻が未入力のとき : 2018/08/10
            //ch = false;
            //for (int i = 0; i < rd.Length; i++)
            //{
            //    if (Utility.StrtoInt(rd[i]) == global.JIYUKBN_CHOKKOU ||
            //        Utility.StrtoInt(rd[i]) == global.JIYUKBN_CHOKKOUCHOKKI ||
            //        Utility.StrtoInt(rd[i]) == global.JIYUKBN_SHUCCHOU)
            //    {
            //        ch = true;
            //        break;
            //    }
            //}

            //if (!ch)
            //{
            //    // 出勤時刻記入なし、退勤時刻記入あり
            //    if (sTimeW == string.Empty && eTimeW != string.Empty)
            //    {
            //        setErrStatus(eSH, iX - 1, "出勤時刻が未入力です");
            //        return false;
            //    }
            //}

            // 記入のとき
            if (m.開始時 != string.Empty || m.開始分 != string.Empty)
            {
                // 数字範囲、単位チェック
                if (!Utility.checkHourSpan(m.開始時, 23))
                {
                    setErrStatus(eSH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!Utility.checkMinSpan(m.開始分, Tani))
                {
                    setErrStatus(eSM, iX - 1, tittle + "が正しくありません");
                    return false;
                }
            }

            if (m.退勤時 != string.Empty || m.退勤分 != string.Empty)
            {
                if (!Utility.checkHourSpan(m.退勤時, 47))
                {
                    setErrStatus(eEH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!Utility.checkMinSpan(m.退勤分, Tani))
                {
                    setErrStatus(eEM, iX - 1, tittle + "が正しくありません");
                    return false;
                }
            }

            if (m.開始時 != string.Empty && m.開始分 != string.Empty &&
                m.退勤時 != string.Empty && m.退勤分 != string.Empty)
            { 
                // 翌日は48時間制で記入する
                if ((Utility.StrtoInt(m.開始時) * 100 + Utility.StrtoInt(m.開始分)) >=
                    (Utility.StrtoInt(m.退勤時) * 100 + Utility.StrtoInt(m.退勤分)))
                {
                    setErrStatus(eSH, iX - 1, "出勤時刻が退勤時刻以後になっています");
                    return false;
                }
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     休憩時間記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckRestTime(CONYX_CLIDataSet.勤務票明細Row m, string tittle, int iX, int Tani)
        {
            // 出退勤時間未記入
            string sTimeW = m.開始時.Trim() + m.開始分.Trim();
            string eTimeW = m.退勤時.Trim() + m.退勤分.Trim();

            if (sTimeW == string.Empty && eTimeW == string.Empty)
            {
                if (m.休憩時 != string.Empty || m.休憩分 != string.Empty)
                {
                    setErrStatus(eRh, iX - 1, "出退勤時刻が未入力で休憩が入力されています");
                    return false;
                }
            }

            // 記入のとき
            if (m.休憩時 != string.Empty || m.休憩分 != string.Empty)
            {
                // 数字範囲、単位チェック
                if (!Utility.checkHourSpan(m.休憩時, 9))
                {
                    setErrStatus(eRh, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!Utility.checkMinSpan(m.休憩分, Tani))
                {
                    setErrStatus(eRm, iX - 1, tittle + "が正しくありません");
                    return false;
                }
            }

            // 出勤～退勤時間
            DateTime stm;
            DateTime etm;

            bool sb = DateTime.TryParse(m.開始時 + ":" + m.開始分, out stm);
            bool ed = DateTime.TryParse(m.退勤時 + ":" + m.退勤分, out etm);
            double rTime = Utility.StrtoDouble(m.休憩時) * 60 + Utility.StrtoDouble(m.休憩分); 

            if (sb && ed)
            {
                double w = Utility.GetTimeSpan(stm, etm).TotalMinutes;
                if (rTime >= w)
                {
                    setErrStatus(eRh, iX - 1, "休憩時間が開始～終業時間以上になっています");
                    return false;
                }
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     実働時間記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckWorkTime(CONYX_CLIDataSet.勤務票明細Row m, string tittle, int iX)
        {
            // 出退勤時間未記入
            string sTimeW = m.開始時.Trim() + m.開始分.Trim();
            string eTimeW = m.退勤時.Trim() + m.退勤分.Trim();

            if (sTimeW == string.Empty && eTimeW == string.Empty)
            {
                if (m.実労時 != string.Empty || m.実労分 != string.Empty)
                {
                    setErrStatus(eWh, iX - 1, "出退勤時刻が未入力で実働時間が入力されています");
                    return false;
                }
            }

            // 出勤時刻、退勤時刻が無記入のときはチェックしない：2018/08/10
            if (sTimeW == string.Empty || eTimeW == string.Empty)
            {
                return true;
            }

            // 出勤～退勤時間
            DateTime stm;
            DateTime etm;

            bool sb = DateTime.TryParse(m.開始時 + ":" + m.開始分, out stm);
            bool ed;
            
            if (Utility.StrtoInt(m.退勤時) > 23)
            {
                // 終業が翌日のとき
                ed = DateTime.TryParse((Utility.StrtoInt(m.退勤時) - 24) + ":" + m.退勤分, out etm);
                etm = etm.AddDays(1);
            }
            else
            {
                // 終業が当日内のとき
                ed = DateTime.TryParse(m.退勤時 + ":" + m.退勤分, out etm);
            }

            if (sb && ed)
            {
                double rTime = Utility.StrtoDouble(m.休憩時) * 60 + Utility.StrtoDouble(m.休憩分);
                double wTime = Utility.StrtoDouble(m.実労時) * 60 + Utility.StrtoDouble(m.実労分);
                double w = Utility.GetTimeSpan(stm, etm).TotalMinutes - rTime;

                if (wTime != w)
                {
                    int wh = (int)(w / 60);
                    int wm = (int)(w % 60);

                    setErrStatus(eWh, iX - 1, "実労時間が終業－開始－休憩（" + wh + ":" + wm.ToString().PadLeft(2, '0') +  "）と一致していません");
                    return false;
                }
            }

            return true;
        }        

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     有給申請チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckYukyuCheck(CONYX_CLIDataSet.勤務票明細Row m, string tittle, int iX)
        {
            //if (m.出勤状況 == global.STATUS_YUKYU)
            //{
            //    if (m.有給申請 == global.flgOff)
            //    {
            //        setErrStatus(eYukyuCheck, iX - 1, "有給申請が未チェックです");
            //        return false;
            //    }
            //}

            return true;
        }
        

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     実労日数チェック </summary>
        /// <param name="r">
        ///     CONYX_CLIDataSet.勤務票ヘッダRow</param>
        /// <param name="mList">
        ///     List<CONYX_CLIDataSet.勤務票明細Row> </param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckWorkDaysTotal(CONYX_CLIDataSet.勤務票ヘッダRow r, List<CONYX_CLIDataSet.勤務票明細Row> mList)
        {
            //int wdays = mList.Count(a => a.出勤状況 == global.STATUS_KIHON_1 || a.出勤状況 == global.STATUS_KIHON_2 || 
            //                        a.出勤状況 == global.STATUS_KIHON_3 || a.出勤時 != string.Empty);

            //if (wdays != r.実労日数)
            //{
            //    setErrStatus(eWorkDays, 0, "実労日数が正しくありません（" + wdays +"日）");
            //    return false;
            //}

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤務実績チェック </summary>
        /// <param name="r">
        ///     CONYX_CLIDataSet.勤務票ヘッダRow</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckNoWork(CONYX_CLIDataSet.勤務票ヘッダRow r)
        {
            //if (r.公休日数 == global.flgOff && r.有休日数 == global.flgOff && r.実労日数 == global.flgOff)
            //{
            //    setErrStatus(eWorkDays, 0, "勤務実績のない出勤簿は登録できません");
            //    return false;
            //}

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     合計日数チェック </summary>
        /// <param name="r">
        ///     CONYX_CLIDataSet.勤務票ヘッダRow</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckTotalDays(CONYX_CLIDataSet.勤務票ヘッダRow r)
        {
            //if ((r.有休日数 + r.実労日数) != r.要出勤日数)
            //{
            //    setErrStatus(eWorkDays, 0, "合計日数が正しくありません（" + (r.有休日数 + r.実労日数) + "日）");
            //    return false;
            //}

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     有給日数チェック </summary>
        /// <param name="r">
        ///     CONYX_CLIDataSet.勤務票ヘッダRow</param>
        /// <param name="mList">
        ///     List<CONYX_CLIDataSet.勤務票明細Row> </param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckYukyuDaysTotal(CONYX_CLIDataSet.勤務票ヘッダRow r, List<CONYX_CLIDataSet.勤務票明細Row> mList)
        {
            //int ydays = mList.Count(a => a.出勤状況 == global.STATUS_YUKYU);

            //if (ydays != r.有休日数)
            //{
            //    setErrStatus(eYukyuDays, 0, "有給日数が正しくありません（" + ydays + "日）");
            //    return false;
            //}

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     公休日数チェック </summary>
        /// <param name="r">
        ///     CONYX_CLIDataSet.勤務票ヘッダRow</param>
        /// <param name="mList">
        ///     List<CONYX_CLIDataSet.勤務票明細Row> </param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckKoukyuDaysTotal(CONYX_CLIDataSet.勤務票ヘッダRow r, List<CONYX_CLIDataSet.勤務票明細Row> mList)
        {
            //int kdays = mList.Count(a => a.出勤状況 == global.STATUS_KOUKYU);

            //if (kdays != r.公休日数)
            //{
            //    setErrStatus(eKoukyuDays, 0, "公休日数が正しくありません（" + kdays + "日）");
            //    return false;
            //}

            return true;
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     検索用DepartmentCodeを取得する </summary>
        /// <returns>
        ///     DepartmentCode</returns>
        ///----------------------------------------------------------
        private string getDepartmentCode(string bCode)
        {
            string strCode = "";

            // DepartmentCode（部署コード）
            if (Utility.NumericCheck(bCode))
            {
                strCode = bCode.PadLeft(15, '0');
            }
            else
            {
                strCode = bCode.PadRight(15, ' ');
            }

            return strCode;
        }
        
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///      勤務票ヘッダデータで指定した社員番号の件数を調べる </summary>
        /// <param name="dts">
        ///     勤務票ヘッダデータセット</param>
        /// <param name="sNum">
        ///     スタッフ番号</param>
        /// <returns>
        ///     件数</returns>
        ///------------------------------------------------------------------------------------
        private bool getSameNumber(CONYX_CLIDataSet dts, int sNum)
        {
            bool rtn = true;

            //if (sNum == string.Empty) return rtn;

            // 指定した社員番号の件数を調べる
            if (dts.勤務票ヘッダ.Count(a => a.社員番号 == sNum) > 1)
            {
                rtn = false;
            }

            return rtn;
        }
        
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     明細記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     行を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckRow(CONYX_CLIDataSet.勤務票明細Row m, string tittle, int iX, string [] st)
        {
            /* 出勤・退勤時刻、残業休日出勤申請、休憩時刻、事由のいずれか記入されて勤務体系が未記入のとき
             * 2018/08/13 */
            if (m.開始時 != string.Empty || m.開始分 != string.Empty ||
                m.退勤時 != string.Empty || m.退勤分 != string.Empty || 
                m.残業休日出勤申請 == global.flgOn || 
                m.休憩時 != string.Empty || m.休憩分 != string.Empty || 
                m.事由１ != string.Empty || m.事由2 != string.Empty)
            {
                if (m.勤務体系コード == string.Empty)
                {
                    setErrStatus(eKinmuTaikeiCode, iX - 1, "勤務体系コードが未入力です");
                    return false;
                }
            }

            bool nk = false;
            for (int i = 0; i < st.Length; i++)
            {
                ///* 事由の取得単位が終日、その他以外のとき */
                //if (Utility.StrtoInt(st[i]) != global.JIYUTANI_ALLDAY &&
                //    Utility.StrtoInt(st[i]) != global.JIYUTANI_255)

                /* 事由の取得単位が終日以外のとき : 2018/08/13 */
                if (Utility.StrtoInt(st[i]) != global.JIYUTANI_ALLDAY)
                {
                    nk = true;
                }
            }
            
            if (nk)
            {
                if (m.開始時 == string.Empty)
                {
                    setErrStatus(eSH, iX - 1, "出勤時刻が未入力です");
                    return false;
                }

                if (m.開始分 == string.Empty)
                {
                    setErrStatus(eSM, iX - 1, "出勤時刻が未入力です");
                    return false;
                }

                if (m.退勤時 == string.Empty)
                {
                    setErrStatus(eEH, iX - 1, "退勤時刻が未入力です");
                    return false;
                }

                if (m.退勤時 == string.Empty)
                {
                    setErrStatus(eEM, iX - 1, "退勤時刻が未入力です");
                    return false;
                }

            }
            
            return true;
        }

        
        
        ///----------------------------------------------------------------------
        /// <summary>
        ///     交通区分チェック </summary>
        /// <param name="m">
        ///     CONYX_CLIDataSet.勤務票明細Row</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///----------------------------------------------------------------------
        private bool errCheckDay(CONYX_CLIDataSet.勤務票ヘッダRow r, CONYX_CLIDataSet.勤務票明細Row m, string tittle, int iX)
        {
            // 日付
            if (m.日 != global.flgOff)
            {
                // 2018/10/02
                DateTime dt;
                int yy = 0;
                int mm = 0;

                if (m.日 < 16)
                {
                    yy = r.年;
                    mm = r.月;
                }
                else
                {
                    if (r.月 == 1)
                    {
                        yy = r.年 - 1;
                        mm = 12; 
                    }
                    else
                    {
                        yy = r.年;
                        mm = r.月 - 1;
                    }
                }

                //string dd = r.年 + "/" + r.月 + "/" + m.日;
                string dd = yy + "/" + mm + "/" + m.日;
                

                //if (!DateTime.TryParse(r.年 + "/" + r.月 + "/" + m.日, out dt))
                if (!DateTime.TryParse(dd, out dt))
                {
                    setErrStatus(eDay, iX - 1, "日付が不正です");
                    return false;
                }
            }

            return true;
        }
        

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間外記入チェック </summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="zan">
        ///     算出残業時間</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckZanTm(CONYX_CLIDataSet.勤務票明細Row m, string tittle, int iX, Int64 zan)
        {
            //Int64 mZan = 0;

            //mZan = (Utility.StrtoInt(m.時間外時) * 60) + Utility.StrtoInt(m.時間外分);

            //// 記入時間と計算された残業時間が不一致のとき
            //if (zan != mZan)
            //{
            //    Int64 hh = zan / 60;
            //    Int64 mm = zan % 60;

            //    setErrStatus(eZH, iX - 1, tittle + "が正しくありません。（" + hh.ToString() + "時間" + mm.ToString() + "分）");
            //    return false;
            //}

            return true;
        }

        /// ----------------------------------------------------------------------------------
        /// <summary>
        ///     時間外算出 2015/09/16 </summary>
        /// <param name="m">
        ///     SCCSDataSet.勤務票明細Row </param>
        /// <param name="Tani">
        ///     丸め単位・分</param>
        /// <param name="ws">
        ///     1日の所定労働時間</param>
        /// <returns>
        ///     時間外・分</returns>
        /// ----------------------------------------------------------------------------------
        public Int64 getZangyoTime(CONYX_CLIDataSet.勤務票明細Row m, Int64 Tani, Int64 ws, Int64 restTime, out Int64 s10Rest, int taikeiCode)
        {
            Int64 zan = 0;  // 計算後時間外勤務時間
            s10Rest = 0;    // 深夜勤務時間帯の10分休憩時間

            DateTime cTm;
            DateTime sTm;
            DateTime eTm;
            DateTime zsTm;
            DateTime pTm;

            //if (!m.Is出勤時Null() && !m.Is出勤分Null() && !m.Is出勤時Null() && !m.Is出勤分Null())
            //{
            //    int ss = Utility.StrtoInt(m.出勤時) * 100 + Utility.StrtoInt(m.出勤分);
            //    int ee = Utility.StrtoInt(m.退勤時) * 100 + Utility.StrtoInt(m.退勤分);
            //    DateTime dt = DateTime.Today;
            //    string sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();

            //    // 始業時刻
            //    if (DateTime.TryParse(sToday + " " + m.出勤時 + ":" + m.出勤分, out cTm))
            //    {
            //        sTm = cTm;
            //    }
            //    else return 0;

            //    // 終業時刻
            //    if (ss > ee)
            //    {
            //        // 翌日
            //        dt = DateTime.Today.AddDays(1);
            //        sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();
            //        if (DateTime.TryParse(sToday + " " + m.退勤時 + ":" + m.退勤分, out cTm))
            //        {
            //            eTm = cTm;
            //        }
            //        else return 0;
            //    }
            //    else
            //    {
            //        // 同日
            //        if (DateTime.TryParse(sToday + " " + m.退勤時 + ":" + m.退勤分, out cTm))
            //        {
            //            eTm = cTm;
            //        }
            //        else return 0;
            //    }


            //    //MessageBox.Show(sTm.ToShortDateString() + " " + sTm.ToShortTimeString() + "    " + eTm.ToShortDateString() + " " + eTm.ToShortTimeString());


            //    // 作業日報に記入されている始業から就業までの就業時間取得
            //    double w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes - restTime;

            //    // 所定労働時間内なら時間外なし
            //    if (w <= ws)
            //    {
            //        return 0;
            //    }

            //    // 所定労働時間＋休憩時間＋10分または15分経過後の時刻を取得（時間外開始時刻）
            //    zsTm = sTm.AddMinutes(ws);          // 所定労働時間
            //    zsTm = zsTm.AddMinutes(restTime);   // 休憩時間
            //    int zSpan = 0;

            //    if (taikeiCode == 100)
            //    {
            //        zsTm = zsTm.AddMinutes(10);         // 体系コード：100 所定労働時間後の10分休憩
            //        zSpan = 130;
            //    }
            //    else if (taikeiCode == 200 || taikeiCode == 300)
            //    {
            //        zsTm = zsTm.AddMinutes(15);         // 体系コード：200,300 所定労働時間後の15分休憩
            //        zSpan = 135;
            //    }
                
            //    pTm = zsTm;                         // 時間外開始時刻

            //    // 該当時刻から終業時刻まで130分または135分以上あればループさせる
            //    while (Utility.GetTimeSpan(pTm, eTm).TotalMinutes > zSpan)
            //    {
            //        // 終業時刻まで2時間につき10分休憩として時間外を算出
            //        // 時間外として2時間加算
            //        zan += 120;

            //        // 130分、または135分後の時刻を取得（2時間＋10分、または15分）
            //        pTm = pTm.AddMinutes(zSpan);

            //        // 深夜勤務時間中の10分または15分休憩時間を取得する
            //        s10Rest += getShinya10Rest(pTm, eTm, zSpan - 120);
            //    }

            //    // 130分（135分）以下の時間外を加算
            //    zan += (Int64)Utility.GetTimeSpan(pTm, eTm).TotalMinutes;

            //    // 単位で丸める
            //    zan -= (zan % Tani);

            //    //MessageBox.Show(pTm.ToShortDateString() + "    " + eTm.ToShortDateString());
            //}
                        
            return zan;
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間中の10分休憩時間を取得する </summary>
        /// <param name="pTm">
        ///     時刻</param>
        /// <param name="eTm">
        ///     終業時刻</param>
        /// <param name="taikeiRest">
        ///     勤務体系別の休憩時間(10分または15分）</param>
        /// <returns>
        ///     休憩時間</returns>
        /// --------------------------------------------------------------------
        private int getShinya10Rest(DateTime pTm, DateTime eTm, int taikeiRest)
        {
            int restTime = 0;

            // 130(135)分後の時刻が終業時刻以内か
            TimeSpan ts = eTm.TimeOfDay;
            
            if (pTm <= eTm)
            {
                // 時刻が深夜時間帯か？
                if (pTm.Hour >= 22 || pTm.Hour <= 5)
                {
                    if (pTm.Hour == 22)
                    {
                        // 22時帯は22時以降の経過分を対象とします。
                        // 例）21:57～22:07のとき22時台の7分が休憩時間
                        if (pTm.Minute >= taikeiRest)
                        {
                            restTime = taikeiRest;
                        }
                        else
                        {
                            restTime = pTm.Minute;
                        }
                    }
                    else if (pTm.Hour == 5)
                    {
                        // 4時帯の経過分を対象とするので5時帯は減算します。
                        // 例）4:57～5:07のとき5時台の7分は差し引いて3分が休憩時間
                        if (pTm.Minute < taikeiRest)
                        {
                            restTime = (taikeiRest - pTm.Minute);
                        }
                    }
                    else
                    {
                        restTime = taikeiRest;
                    }
                }
            }

            return restTime;
        }


        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務記入チェック </summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="Tani">
        ///     分記入単位</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckShinya(CONYX_CLIDataSet.勤務票明細Row m, string tittle, int Tani, int iX)
        {
        //    // 無記入なら終了
        //    if (m.深夜時 == string.Empty && m.深夜分 == string.Empty) return true;

        //    //  始業、終業時刻が無記入で深夜が記入されているときエラー
        //    if (m.開始時 == string.Empty && m.開始分 == string.Empty &&
        //         m.終了時 == string.Empty && m.終了分 == string.Empty)
        //    {
        //        if (m.深夜時 != string.Empty)
        //        {
        //            setErrStatus(eSIH, iX - 1, "始業、終業時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }

        //        if (m.深夜分 != string.Empty)
        //        {
        //            setErrStatus(eSIM, iX - 1, "始業、終業時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }
        //    }

        //    // 記入のとき
        //    if (m.深夜時 != string.Empty || m.深夜分 != string.Empty)
        //    {
        //        // 時間と分のチェック
        //        //if (!checkHourSpan(m.時間外時))
        //        //{
        //        //    setErrStatus(eZH, iX - 1, tittle + "が正しくありません");
        //        //    return false;
        //        //}

        //        if (!checkMinSpan(m.深夜分, Tani))
        //        {
        //            setErrStatus(eSIM, iX - 1, tittle + "が正しくありません。（" + Tani.ToString() + "分単位）");
        //            return false;
        //        }
        //    }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間チェック </summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="shinya">
        ///     算出された深夜k勤務時間</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckShinyaTm(CONYX_CLIDataSet.勤務票明細Row m, string tittle, int iX, Int64 shinya)
        {
            Int64 mShinya = 0;

            //mShinya = (Utility.StrtoInt(m.深夜時) * 60) + Utility.StrtoInt(m.深夜分);

            //// 記入時間と計算された深夜時間が不一致のとき
            //if (shinya != mShinya)
            //{
            //    Int64 hh = shinya / 60;
            //    Int64 mm = shinya % 60;

            //    setErrStatus(eSIH, iX - 1, tittle + "が正しくありません。（" + hh.ToString() + "時間" + mm.ToString() + "分）");
            //    return false;
            //}

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     実働時間を取得する</summary>
        /// <param name="sH">
        ///     開始時</param>
        /// <param name="sM">
        ///     開始分</param>
        /// <param name="eH">
        ///     終了時</param>
        /// <param name="eM">
        ///     終了分</param>
        /// <param name="rH">
        ///     休憩時間・分</param>
        /// <returns>
        ///     実働時間</returns>
        ///------------------------------------------------------------------------------------
        public double getWorkTime(string sH, string sM, string eH, string eM, int rH)
        {
            DateTime sTm;
            DateTime eTm;
            DateTime cTm;
            double w = 0;   // 稼働時間

            // 時刻情報に不備がある場合は０を返す
            if (!Utility.NumericCheck(sH) || !Utility.NumericCheck(sM) || 
                !Utility.NumericCheck(eH) || !Utility.NumericCheck(eM))
                return 0;

            int ss = Utility.StrtoInt(sH) * 100 + Utility.StrtoInt(sM);
            int ee = Utility.StrtoInt(eH) * 100 + Utility.StrtoInt(eM);
            DateTime dt = DateTime.Today;
            string sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();

            // 開始時刻取得
            if (Utility.StrtoInt(sH) == 24)
            {
                if (DateTime.TryParse(sToday + " 0:" + Utility.StrtoInt(sM).ToString(), out cTm))
                {
                    sTm = cTm;
                }
                else return 0;
            }
            else
            {
                if (DateTime.TryParse(sToday + " " + Utility.StrtoInt(sH).ToString() + ":" + Utility.StrtoInt(sM).ToString(), out cTm))
                {
                    sTm = cTm;
                }
                else return 0;
            }
            
            // 終業時刻
            if (ss > ee)
            {
                // 翌日
                dt = DateTime.Today.AddDays(1);
                sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();
            }

            // 終了時刻取得
            if (Utility.StrtoInt(eH) == 24)
                eTm = DateTime.Parse(sToday + " 23:59");
            else
            {
                if (DateTime.TryParse(sToday + " " + Utility.StrtoInt(eH).ToString() + ":" + Utility.StrtoInt(eM).ToString(), out cTm))
                {
                    eTm = cTm;
                }
                else return 0;
            }

            // 終了時間が24:00記入のときは23:59までの計算なので稼働時間1分加算する
            if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
            {
                w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes + 1;
            }
            else if (sTm == eTm)    // 同時刻の場合は翌日の同時刻とみなす 2014/10/10
            {
                w = Utility.GetTimeSpan(sTm, eTm.AddDays(1)).TotalMinutes;  // 稼働時間
            }
            else
            {
                w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes;  // 稼働時間
            }

            // 休憩時間を差し引く
            if (w >= rH) w = w - rH;
            else w = 0;

            // 値を返す
            return w;
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間を取得する</summary>
        /// <param name="sH">
        ///     開始時</param>
        /// <param name="sM">
        ///     開始分</param>
        /// <param name="eH">
        ///     終了時</param>
        /// <param name="eM">
        ///     終了分</param>
        /// <param name="tani">
        ///     丸め単位</param>
        /// <param name="s10">
        ///     深夜勤務時間中の10分休憩</param>
        /// <returns>
        ///     深夜勤務時間・分</returns>
        /// ------------------------------------------------------------
        public double getShinyaWorkTime(string sH, string sM, string eH, string eM, int tani, Int64 s10)
        {
            DateTime sTime;
            DateTime eTime;
            DateTime cTm;

            double wkShinya = 0;    // 深夜稼働時間

            // 時刻情報に不備がある場合は０を返す
            if (!Utility.NumericCheck(sH) || !Utility.NumericCheck(sM) ||
                !Utility.NumericCheck(eH) || !Utility.NumericCheck(eM))
                return 0;

            // 開始時間を取得
            if (DateTime.TryParse(Utility.StrtoInt(sH).ToString() + ":" + Utility.StrtoInt(sM).ToString(), out cTm))
            {
                sTime = cTm;
            }
            else return 0;

            // 終了時間を取得
            if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
            {
                eTime = global.dt2359;
            }
            else if (DateTime.TryParse(Utility.StrtoInt(eH).ToString() + ":" + Utility.StrtoInt(eM).ToString(), out cTm))
            {
                eTime = cTm;
            }
            else return 0;


            // 当日内の勤務のとき
            if (sTime.TimeOfDay < eTime.TimeOfDay)
            {
                // 早出残業時間を求める
                if (sTime < global.dt0500)  // 開始時刻が午前5時前のとき
                {
                    // 早朝時間帯稼働時間
                    if (eTime >= global.dt0500)
                    {
                        wkShinya += Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;
                    }
                    else
                    {
                        wkShinya += Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
                    }
                }

                // 終了時刻が22:00以降のとき
                if (eTime >= global.dt2200)
                {
                    // 当日分の深夜帯稼働時間を求める
                    if (sTime <= global.dt2200)
                    {
                        // 出勤時刻が22:00以前のとき深夜開始時刻は22:00とする
                        wkShinya += Utility.GetTimeSpan(global.dt2200, eTime).TotalMinutes;
                    }
                    else
                    {
                        // 出勤時刻が22:00以降のとき深夜開始時刻は出勤時刻とする
                        wkShinya += Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
                    }

                    // 終了時間が24:00記入のときは23:59までの計算なので稼働時間1分加算する
                    if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
                        wkShinya += 1;
                }
            }
            else
            {
                // 日付を超えて終了したとき（開始時刻 >= 終了時刻）※2014/10/10 同時刻は翌日の同時刻とみなす

                // 早出残業時間を求める
                if (sTime < global.dt0500)  // 開始時刻が午前5時前のとき
                {
                    wkShinya += Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;
                }

                // 当日分の深夜勤務時間（～０：００まで）
                if (sTime <= global.dt2200)
                {
                    // 出勤時刻が22:00以前のとき無条件に120分
                    wkShinya += global.TOUJITSU_SINYATIME;
                }
                else
                {
                    // 出勤時刻が22:00以降のとき出勤時刻から24:00までを求める
                    wkShinya += Utility.GetTimeSpan(sTime, global.dt2359).TotalMinutes + 1;
                }

                // 0:00以降の深夜勤務時間を加算（０：００～終了時刻）
                if (eTime.TimeOfDay > global.dt0500.TimeOfDay)
                {
                    wkShinya += Utility.GetTimeSpan(global.dt0000, global.dt0500).TotalMinutes;
                }
                else
                {
                    wkShinya += Utility.GetTimeSpan(global.dt0000, eTime).TotalMinutes;
                }
            }

            // 深夜勤務時間中の10分または15分休憩時間を差し引く
            wkShinya -= s10;

            // 単位分で丸め
            wkShinya -= (wkShinya % tani);

            return wkShinya;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     社員コードチェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl オブジェクト </param>
        /// <param name="s">
        ///     社員番号</param>
        /// <param name="sDt">
        ///     基準年月日</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        private bool chkShainCode(string s, sqlControl.DataControl sdCon)
        {
            bool dm = false;

            // 社員コード取得
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select EmployeeNo,RetireCorpDate from tbEmployeeBase ");
            sb.Append("where EmployeeNo = '" + s.PadLeft(10, '0') + "' ");
            //sb.Append(" and BeOnTheRegisterDivisionID != 9");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                dm = true;
                break;
            }

            dR.Close();

            return dm;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     勤務体系コード存在チェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl オブジェクト </param>
        /// <param name="s">
        ///     勤務体系コード</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        private bool chkKinmuTaikeiCode(CONYX_CLIDataSet.勤務票明細Row m, string tittle, int iX, sqlControl.DataControl sdCon)
        {
            if (m.勤務体系コード == string.Empty)
            {
                return true;
            }

            // 勤務体系コード
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select LaborSystemCode, LaborSystemName from tbLaborSystem ");
            sb.Append("where LaborSystemCode = '" + m.勤務体系コード.PadLeft(4, '0') + "'");
            
            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());
            bool dm = dR.HasRows;

            dR.Close();
            
            if (!dm)
            {
                setErrStatus(eKinmuTaikeiCode, iX - 1, tittle + "が正しくありません");
                return false;
            }

            return true;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     事由コード存在チェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl オブジェクト </param>
        /// <param name="s">
        ///     事由コード</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        private bool chkJiyuCode(string jiyu, string tittle, int iX, sqlControl.DataControl sdCon, int errNum)
        {
            // 事由コード
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select LaborReasonCode from tbLaborReason ");
            sb.Append("where LaborReasonCode = '" + jiyu.PadLeft(2, '0') + "'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());
            bool dm = dR.HasRows;

            dR.Close();

            if (!dm)
            {
                setErrStatus(errNum, iX - 1, tittle + "が正しくありません");
                return false;
            }

            return true;
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの普通残業時間に該当するか調べる </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="m">
        ///     CONYX_CLIDataSet.勤務票明細Row</param>
        ///----------------------------------------------------------------------------------
        private bool isZangyo(sqlControl.DataControl sdCon, CONYX_CLIDataSet.勤務票明細Row m)
        {
            // 出勤時間空白は対象外とする
            if (m.開始時 == string.Empty || m.開始分 == string.Empty ||
                m.退勤時 == string.Empty || m.退勤分 == string.Empty)
            {
                return false;
            }

            bool bn = false;
            DateTime dtStart = DateTime.Now;        // 普通残業開始時刻
            DateTime dtEnd = DateTime.Now;          // 普通残業終了時刻
            DateTime dtChange = DateTime.Now;       // 日替わり時刻
                        
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select tbLaborSystem.LaborSystemCode, LaborSystemName, tbLaborSystem.LatterHalfStartTime, tbLaborSystem.FirstHalfEndTime,");
            sb.Append("tbLaborSystem.DayChangeTime,a.StartTime,a.EndTime, AttendanceType ");
            sb.Append("from tbLaborSystem left join ");
            sb.Append("(select * from tbLaborTimeSpanRule where LaborTimeItemID = 4) as a ");
            sb.Append("on tbLaborSystem.LaborSystemID = a.LaborSystemID ");
            sb.Append("where tbLaborSystem.LaborSystemCode = '").Append(m.勤務体系コード).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                // 休日出勤のとき
                if (Utility.StrtoInt(dR["AttendanceType"].ToString()) == 1 ||
                    Utility.StrtoInt(dR["AttendanceType"].ToString()) == 2)
                {
                    dR.Close();
                    return true;
                }

                if (!(dR["StartTime"] is DBNull))
                {
                    bn = true;
                    dtStart = (DateTime)dR["StartTime"];
                    dtChange = (DateTime)dR["DayChangeTime"];
                }

                //if (!(dR["EndTime"] is DBNull))
                //{
                //    bn = true;
                //    dtEnd = (DateTime)dR["EndTime"];
                //    dtChange = (DateTime)dR["DayChangeTime"];
                //}
            }

            dR.Close();

            // 残業開始時刻がないときは戻る
            if (!bn)
            {
                return false;
            }

            // 普通残業開始時間の当日・翌日判定
            if (dtStart.Day == 2)
            {
                // 当日日付
                dtStart = DateTime.Parse(dtStart.Hour.ToString() + ":" + dtStart.Minute.ToString() + ":0");
            }
            else if (dtStart.Day == 3)
            {
                // 翌日日付
                dtStart = DateTime.Parse(dtStart.Hour.ToString() + ":" + dtStart.Minute.ToString() + ":0").AddDays(1);
            }   
            
            DateTime wStartTm = DateTime.Parse(m.開始時 + ":" + m.開始分 + ":0");
            DateTime wEndTm = DateTime.Today; // 仮の初期値

            // 退勤時刻
            int eh = 0;
            if (Utility.StrtoInt(m.退勤時) > 23)
            {
                eh = Utility.StrtoInt(m.退勤時) - 24;
            }
            else
            {
                eh = Utility.StrtoInt(m.退勤時);
            }

            wEndTm = DateTime.Parse(eh + ":" + m.退勤分 + ":0");

            // 日替わり時刻
            int nextTime = dtChange.Hour * 100 + dtChange.Minute;

            int eTm = eh * 100 + Utility.StrtoInt(m.退勤分);
            if (eTm < nextTime)
            {
                // 勤務終了時刻が日替わり時刻以前のとき
                wEndTm = wEndTm.AddDays(1);  // 日付を翌日にする
            }


            //debug
            System.Diagnostics.Debug.WriteLine("日替わり時刻:" + nextTime + "eTm:" + eTm + " 退勤時刻：" + wEndTm + " 普通残業開始時刻：" + dtStart);
            
                        
            // 退勤時刻が普通残業開始時刻以降のとき超過時間計算を行う
            if (wEndTm > dtStart)
            {
                return true;
            }

            //// 出勤時刻が勤務開始時刻以前のとき早出時間計算を行う
            //if (wStartTm < dtStart)
            //{
            //    //if (dtStart < wEndTm)
            //    //{
            //    //    /* 退勤時刻時刻が勤務体系の開始時刻以降のとき
            //    //     * 例）勤務体系8:00-17:00で勤務時間が6:00-17:00等 */
            //    //    zanTime += Utility.GetTimeSpan(wStartTm, dtStart).TotalMinutes;
            //    //}
            //    //else
            //    //{
            //    //    /* 退勤時刻時刻が勤務体系の開始時刻以前のとき
            //    //     * 例）勤務体系8:00-17:00で勤務時間が6:00-7:00等 */
            //    //    zanTime += Utility.GetTimeSpan(wStartTm, wEndTm).TotalMinutes;
            //    //}

            //    return true;
            //}

            return false;
        }


        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの早出残業時間に該当するか調べる </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="m">
        ///     CONYX_CLIDataSet.勤務票明細Row</param>
        ///----------------------------------------------------------------------------------
        private bool isHayade(sqlControl.DataControl sdCon, CONYX_CLIDataSet.勤務票明細Row m)
        {
            // 出勤時間空白は対象外とする
            if (m.開始時 == string.Empty || m.開始分 == string.Empty ||
                m.退勤時 == string.Empty || m.退勤分 == string.Empty)
            {
                return false;
            }

            bool bn = false;
            DateTime dtStart = DateTime.Now;        // 普通残業開始時刻
            DateTime dtEnd = DateTime.Now;          // 普通残業終了時刻
            DateTime dtChange = DateTime.Now;       // 日替わり時刻

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select tbLaborSystem.LaborSystemCode, LaborSystemName, tbLaborSystem.LatterHalfStartTime, tbLaborSystem.FirstHalfEndTime,");
            sb.Append("tbLaborSystem.DayChangeTime,a.StartTime,a.EndTime, AttendanceType ");
            sb.Append("from tbLaborSystem left join ");
            sb.Append("(select * from tbLaborTimeSpanRule where LaborTimeItemID = 7) as a ");
            sb.Append("on tbLaborSystem.LaborSystemID = a.LaborSystemID ");
            sb.Append("where tbLaborSystem.LaborSystemCode = '").Append(m.勤務体系コード).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                //// 休日出勤のとき
                //if (Utility.StrtoInt(dR["AttendanceType"].ToString()) == 1 ||
                //    Utility.StrtoInt(dR["AttendanceType"].ToString()) == 2)
                //{
                //    dR.Close();
                //    return true;
                //}

                if (!(dR["StartTime"] is DBNull))
                {
                    bn = true;
                    dtStart = (DateTime)dR["StartTime"];
                    dtChange = (DateTime)dR["DayChangeTime"];
                }

                if (!(dR["EndTime"] is DBNull))
                {
                    bn = true;
                    dtEnd = (DateTime)dR["EndTime"];
                    dtChange = (DateTime)dR["DayChangeTime"];
                }
            }

            dR.Close();

            // 残業開始時刻がないときは戻る
            if (!bn)
            {
                return false;
            }

            // 早出残業開始時刻の当日・翌日判定
            if (dtStart.Day == 2)
            {
                // 当日日付
                dtStart = DateTime.Parse(dtStart.Hour.ToString() + ":" + dtStart.Minute.ToString() + ":0");
            }
            else if (dtStart.Day == 3)
            {
                // 翌日日付
                dtStart = DateTime.Parse(dtStart.Hour.ToString() + ":" + dtStart.Minute.ToString() + ":0").AddDays(1);
            }

            // 早出残業終了時刻の当日・翌日判定
            if (dtEnd.Day == 2)
            {
                // 当日日付
                dtEnd = DateTime.Parse(dtEnd.Hour.ToString() + ":" + dtEnd.Minute.ToString() + ":0");
            }
            else if (dtEnd.Day == 3)
            {
                // 翌日日付
                dtEnd = DateTime.Parse(dtEnd.Hour.ToString() + ":" + dtEnd.Minute.ToString() + ":0").AddDays(1);
            }

            DateTime wStartTm = DateTime.Today; // 仮の初期値
            DateTime wEndTm = DateTime.Today; // 仮の初期値

            // 出勤時刻
            int sh = 0;
            if (Utility.StrtoInt(m.開始時) > 23)
            {
                sh = Utility.StrtoInt(m.開始時) - 24;
            }
            else
            {
                sh = Utility.StrtoInt(m.開始時);
            }

            wStartTm = DateTime.Parse(sh + ":" + m.開始分 + ":0");

            // 退勤時刻
            int eh = 0;
            if (Utility.StrtoInt(m.退勤時) > 23)
            {
                eh = Utility.StrtoInt(m.退勤時) - 24;
            }
            else
            {
                eh = Utility.StrtoInt(m.退勤時);
            }

            wEndTm = DateTime.Parse(eh + ":" + m.退勤分 + ":0");

            // 日替わり時刻
            int nextTime = dtChange.Hour * 100 + dtChange.Minute;

            int sTm = sh * 100 + Utility.StrtoInt(m.開始分);
            //if (sTm < nextTime)
            //{
            //    // 勤務開始時刻が日替わり時刻以前のとき
            //    wStartTm = wStartTm.AddDays(1);  // 日付を翌日にする
            //}

            int eTm = eh * 100 + Utility.StrtoInt(m.退勤分);
            if (eTm < nextTime)
            {
                // 勤務終了時刻が日替わり時刻以前のとき
                wEndTm = wEndTm.AddDays(1);  // 日付を翌日にする
            }


            //debug
            System.Diagnostics.Debug.WriteLine("日替わり時刻:" + nextTime + " sTm:" + sTm + " eTm:" + eTm + " 出勤時刻：" + wStartTm + " 退勤時刻：" + wEndTm + " 早出開始：" + dtStart + " 早出終了：" + dtEnd);


            // 出勤時刻が早出終了時刻以前のとき早出とみなす
            if (wStartTm < dtEnd)
            {
                return true;
            }

            //// 出勤時刻が勤務開始時刻以前のとき早出時間計算を行う
            //if (wStartTm < dtStart)
            //{
            //    //if (dtStart < wEndTm)
            //    //{
            //    //    /* 退勤時刻時刻が勤務体系の開始時刻以降のとき
            //    //     * 例）勤務体系8:00-17:00で勤務時間が6:00-17:00等 */
            //    //    zanTime += Utility.GetTimeSpan(wStartTm, dtStart).TotalMinutes;
            //    //}
            //    //else
            //    //{
            //    //    /* 退勤時刻時刻が勤務体系の開始時刻以前のとき
            //    //     * 例）勤務体系8:00-17:00で勤務時間が6:00-7:00等 */
            //    //    zanTime += Utility.GetTimeSpan(wStartTm, wEndTm).TotalMinutes;
            //    //}

            //    return true;
            //}

            return false;
        }

    }
}
