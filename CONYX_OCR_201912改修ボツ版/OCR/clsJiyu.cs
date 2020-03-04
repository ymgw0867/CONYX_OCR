using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using CONYX_OCR.common;

namespace CONYX_OCR.OCR
{
    /// <summary>
    ///     事由エラーチェック基本クラス
    /// </summary>
    public class clsJiyu
    {
        public clsJiyu(string[] s)
        {
            // 事由配列
            mJiyu = s;
        }

        protected string[] mJiyu = null;

        //internal sqlControl.DataControl sdCon = null;

        protected SqlDataReader getLaborReason(string sCode, sqlControl.DataControl sdCon)
        {
            //// 奉行SQLServer接続文字列取得
            //string sc = sqlControl.obcConnectSting.get(dbName);
            //sdCon = new sqlControl.DataControl(sc);

            // 事由データ取得
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select LaborReasonCode,AcquireUnit,AcquireDivision,LaborReasonDivision from tbLaborReason ");
            sb.Append("where IsValid = 1 and LaborReasonCode = '" + sCode.PadLeft(2, '0') + "'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            return dR;
        }
    }

    /// <summary>
    ///     事由コード登録チェッククラス：基本クラスを継承</summary>
    public class clsJiyuHas : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     事由コードチェッククラス</summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="db">
        ///     データベース名</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public clsJiyuHas(string[] s)
            : base(s)
        {

        }

        ///------------------------------------------------------------
        /// <summary>
        ///     事由コードチェック </summary>
        /// <param name="errNum">
        ///     エラー事由番号</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isHasRows(out int errNum, sqlControl.DataControl sdCon)
        {
            bool dm = true;
            errNum = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                // 事由データ取得
                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                if (!dR.HasRows)
                {
                    dm = false;
                    errNum = i;
                    dR.Close();
                    break;
                }

                dR.Close();
            }

            //if (sdCon != null)
            //{
            //    sdCon.Close();
            //}

            return dm;
        }
    }

    /// <summary>
    ///     「終日」事由と他の事由併記チェッククラス：基本クラスを継承
    /// </summary>
    public class clsJiyuAllDay : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由と他の事由併記チェッククラス </summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="_db">
        ///     データベース名</param>
        ///------------------------------------------------------------
        public clsJiyuAllDay(string[] s)
            : base(s)
        {

        }

        bool dm = false;

        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由と他の事由併記チェック </summary>
        /// <param name="errNum">
        ///     エラー事由番号</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControlオブジェクト</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isAllDayAnotherDay(out int errNum, sqlControl.DataControl sdCon)
        {
            errNum = 0;
            int jCnt = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得区分
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
                    {
                        // 終日「０」
                        dm = true;  // 終日あり
                        jCnt++;
                    }
                    else if (Utility.NulltoStr(dR["AcquireUnit"]) != "255")
                    {
                        jCnt++;
                    }

                    break;
                }

                dR.Close();
            }

            //if (sdCon != null)
            //{
            //    sdCon.Close();
            //}

            // 終日事由があり、他の事由が併記されているときはエラ―
            if (!dm)
            {
                // 終日がない場合戻る
                return true;
            }
            else
            {
                //int cnt = 0;
                //for (int i = 0; i < mJiyu.Length; i++)
                //{
                //    if (mJiyu[i] != string.Empty)
                //    {
                //        // 事由記入あり
                //        cnt++;
                //        errNum = i;
                //    }
                //}

                //// 終日を含んで2つ以上の事由が記入されているとエラー
                //if (cnt > 1)
                //{
                //    return false;
                //}
                //else
                //{
                //    return true;
                //}

                if (jCnt > 1)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }
    }

    /// <summary>
    ///     「終日」事由とシフト以外の記入チェッククラス：基本クラスを継承
    /// </summary>
    public class clsAlldayAnotherData : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由とシフト以外の記入チェック </summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="_dbName">
        ///     データベース名</param>
        ///------------------------------------------------------------
        public clsAlldayAnotherData(string[] s)
            : base(s)
        {

        }

        bool dm = false;

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     事由の事由区分を取得する </summary>
        /// <param name="m">
        ///     CONYX_CLIDataSet.勤務票明細Row</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <returns>
        ///     事由区分配列</returns>
        ///-----------------------------------------------------------------------
        public string [] getLaborReasonDivision(CONYX_CLIDataSet.勤務票明細Row m, sqlControl.DataControl sdCon)
        {
            string[] rd = new string[2];

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    rd[i] = string.Empty;
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 事由区分を取得：2018/08/10
                    rd[i] = Utility.NulltoStr(dR["LaborReasonDivision"]);
                    break;
                }

                dR.Close();
            }

            return rd;
        }


        ///-----------------------------------------------------------------------
        /// <summary>
        ///     事由の取得単位を取得する </summary>
        /// <param name="m">
        ///     CONYX_CLIDataSet.勤務票明細Row</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <returns>
        ///     取得単位配列</returns>
        ///-----------------------------------------------------------------------
        public string[] getAcquireUnit(CONYX_CLIDataSet.勤務票明細Row m, sqlControl.DataControl sdCon)
        {
            string[] rd = new string[2];

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    rd[i] = string.Empty;
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得単位を取得：2018/08/10
                    rd[i] = Utility.NulltoStr(dR["AcquireUnit"]);
                    break;
                }

                dR.Close();
            }

            return rd;
        }


        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由とシフト以外の記入チェック </summary>
        /// <param name="m"> 
        ///     CONYX_CLIDataSet.勤務票明細Row </param>
        /// <param name="sdCon">
        ///     sqlControl.DataControlオブジェクト</param>
        /// <param name="eNum">
        ///     エラー項目番号</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isAlldayAnotherData(CONYX_CLIDataSet.勤務票明細Row m, sqlControl.DataControl sdCon, out int eNum)
        {
            eNum = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得区分
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
                    {
                        // 終日「０」
                        dm = true;  // 終日あり
                    }

                    break;
                }

                dR.Close();
            }

            //if (sdCon != null)
            //{
            //    sdCon.Close();
            //}

            if (!dm)
            {
                // 終日がない場合戻る
                return true;
            }
            else
            {
                // 終日でシフトコード以外記入があるとき
                if (m.開始時 != string.Empty)
                {
                    eNum = 1;
                    return false;
                }
                if (m.開始分 != string.Empty)
                {
                    eNum = 2;
                    return false;
                }

                if (m.退勤時 != string.Empty)
                {
                    eNum = 3;
                    return false;
                }

                if (m.退勤分 != string.Empty)
                {
                    eNum = 4;
                    return false;
                }

                if (m.休憩時 != string.Empty)
                {
                    eNum = 5;
                    return false;
                }

                if (m.休憩分 != string.Empty)
                {
                    eNum = 6;
                    return false;
                }

                if (m.実労時 != string.Empty)
                {
                    eNum = 7;
                    return false;
                }

                if (m.実労分 != string.Empty)
                {
                    eNum = 8;
                    return false;
                }
                
                return true;
            }
        }
    }
    
    /// <summary>
    ///     取得単位「半日」事由の取得区分の重複記入チェッククラス：基本クラスを継承
    /// </summary>
    public class clsJiyuDiv : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     取得単位「半日」事由の取得区分の重複記入チェッククラス </summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="_dbName">
        ///     データベース名</param>
        ///------------------------------------------------------------
        public clsJiyuDiv(string[] s)
            : base(s)
        {

        }

        bool dm = false;

        ///------------------------------------------------------------
        /// <summary>
        ///     取得単位「半日」事由の取得区分の重複記入チェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControlオブジェクト</param>
        /// <param name="dCnt">
        ///     半日事由の数</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isJiyuDiv(sqlControl.DataControl sdCon, out int dCnt)
        {
            dCnt = 0;            
            string[] div = new string[2];

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    div[i] = string.Empty;
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得単位
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGON)
                    {
                        // 半日「1」
                        dm = true;  // 半日あり

                        // 取得区分 0:前半, 1:後半
                        div[i] = Utility.NulltoStr(dR["AcquireDivision"]); 
                    }

                    break;
                }

                dR.Close();
            }

            //if (sdCon != null)
            //{
            //    sdCon.Close();
            //}

            if (!dm)
            {
                // 半日がない場合戻る
                return true;
            }
            else
            {
                int kbn = 0;

                for (int i = 0; i < 2; i++)
                {
                    if (Utility.NulltoStr(div[i]) == string.Empty)
                    {
                        continue;
                    }

                    kbn += Utility.StrtoInt(div[i]);
                    dCnt++;
                }

                if (dCnt == 3)
                {
                    // 半日事由が3つ記入されている
                    return false;
                }
                else if (dCnt == 2)
                {
                    // 半日事由が2つ記入されている
                    if (kbn != 1)
                    {
                        // 前半(0)・前半(0)、または後半(1)・後半(1)の組み合わせになっている
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return true;
                }
            }
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     記入されている事由が終日単位か半日単位か調べる : 2017/11/21</summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <returns>
        ///     終日は1, 半日は0.5</returns>
        ///----------------------------------------------------------------------------------
        public double jiyuDivCount(sqlControl.DataControl sdCon)
        {
            double dCnt = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i] == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得単位
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
                    {
                        // 終日事由
                        dCnt++;
                    }
                    else if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGON)
                    {
                        // 半日事由
                        dCnt += 0.5;
                    }

                    break;
                }

                dR.Close();
            }

            return dCnt;
        }
    }


    /// <summary>
    ///     「終日」事由とシフト以外の記入チェッククラス：基本クラスを継承
    /// </summary>
    public class clsHanDivShift : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由とシフト以外の記入チェック </summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="_dbName">
        ///     データベース名</param>
        ///------------------------------------------------------------
        public clsHanDivShift(string[] s)
            : base(s)
        {

        }

        bool dm = false;

        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」以外の事由で勤務体系の記入チェック </summary>
        /// <param name="m"> 
        ///     CONYX_CLIDataSet.勤務票明細Row </param>
        /// <param name="sdCon">
        ///     sqlControl.DataControlオブジェクト</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isHanDivShift(CONYX_CLIDataSet.勤務票明細Row m, sqlControl.DataControl sdCon)
        {
            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得区分
                    if (Utility.NulltoStr(dR["AcquireUnit"]) != global.FLGOFF)
                    {
                        // 終日以外あり
                        dm = true;
                    }

                    break;
                }

                dR.Close();
            }

            if (!dm)
            {
                // 該当しない場合戻る
                return true;
            }
            else if (m.勤務体系コード == string.Empty)
            {
                // 勤務体系が未記入のとき
                return false;
            }
            else
            { 
                return true;
            }
        }
    }
}
