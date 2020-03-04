using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Data.OleDb;
using CONYX_OCR.common;

namespace CONYX_OCR.OCR
{
    partial class frmPastCorrect
    {
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤務票ヘッダと勤務票明細のデータセットにデータを読み込む </summary>
        ///------------------------------------------------------------------------------------
        private void getDataSet(string sID)
        {
            adpMn.過去勤務票ヘッダTableAdapter.FillByID(dts.過去勤務票ヘッダ, sID);
            adpMn.過去勤務票明細TableAdapter.FillByHDID(dts.過去勤務票明細, sID);

            // 編集ログデータ
            gAdp.FillByHDID(dts.出勤簿編集ログ, sID);
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     データを画面に表示します </summary>
        /// <param name="iX">
        ///     ヘッダデータID</param>
        ///------------------------------------------------------------------------------------
        private void showOcrData(string iX)
        {
            // 非ログ書き込み状態とする
            editLogStatus = false;

            // 勤務票ヘッダテーブル行を取得
            CONYXDataSet.過去勤務票ヘッダRow r = (CONYXDataSet.過去勤務票ヘッダRow)dts.過去勤務票ヘッダ.Single(a => a.ID == iX);

            // フォーム初期化
            formInitialize(dID);
            
            txtYear.Text = r.年.ToString();
            txtMonth.Text = Utility.EmptytoZero(r.月.ToString());
            txtCode.Text = r.社員番号.ToString();
            lblName.Text = r.社員名;

            if (r.確認 == global.flgOn)
            {
                chkKakunin.Checked = true;
            }
            else
            {
                chkKakunin.Checked = false;
            }

            txtMemo.Text = r.備考;

            // 編集済み項目のバックカラー処理
            foreach (var item in dts.出勤簿編集ログ.Where(a => a.勤務表ヘッダID == r.ID))
            {
                if (item.項目名 == LOG_YEAR)
                {
                    if (global.cnfEditBackColor != "")
                    {
                        txtYear.BackColor = Color.FromArgb(Convert.ToInt32(global.cnfEditBackColor, 16));
                    }                    
                }

                if (item.項目名 == LOG_MONTH)
                {
                    if (global.cnfEditBackColor != "")
                    {
                        txtMonth.BackColor = Color.FromArgb(Convert.ToInt32(global.cnfEditBackColor, 16));
                    }
                }

                if (item.項目名 == LOG_NUMBER)
                {
                    if (global.cnfEditBackColor != "")
                    {
                        txtCode.BackColor = Color.FromArgb(Convert.ToInt32(global.cnfEditBackColor, 16));
                    }
                }
            }
            
            // 日別勤怠表示
            showItem(r.ID, dGV, r.年.ToString(), r.月.ToString());
     
            // エラー情報表示初期化
            lblErrMsg.Visible = false;
            lblErrMsg.Text = string.Empty;

            // 画像フォルダパス
            string dirName = Properties.Settings.Default.tifPath + r.画像名.Substring(0, 6) + @"\";

            // 画像表示 
            ShowImage(dirName + r.画像名.ToString());

            // ログ書き込み状態とする
            editLogStatus = true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤怠明細表示 </summary>
        /// <param name="sID">
        ///     帳票ID １：本社, ２：静岡, ３：大阪製造部</param>
        /// <param name="hID">
        ///     ヘッダID</param>
        /// <param name="sYY">
        ///     年</param>
        /// <param name="sMM">
        ///     月</param>
        /// <param name="dGV">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------------------------------
        private void showItem(string hID, DataGridView dGV, string sYY, string sMM)
        {
            // 社員別勤務実績表示
            int mC = dts.過去勤務票明細.Count(a => a.ヘッダID == hID);
                
            // 行数を設定して表示色を初期化
            dGV.Rows.Clear();
            dGV.RowCount = mC;

            for (int i = 0; i < mC; i++)
            {
                dGV.Rows[i].DefaultCellStyle.BackColor = Color.FromName("Control");
                dGV.Rows[i].ReadOnly = true;    // 初期設定は編集不可とする
            }
                        
            // 行インデックス初期化
            int mRow = 0;

            foreach (var t in dts.過去勤務票明細.Where(a => a.ヘッダID == hID).OrderBy(a => a.ID))
            {
                // 表示色を初期化
                //dGV.Rows[mRow].DefaultCellStyle.BackColor = Color.Empty;
                dGV.Rows[mRow].DefaultCellStyle.BackColor = Color.FromArgb(255, 254, 255);

                // 編集を可能とする
                dGV.Rows[mRow].ReadOnly = false;

                //DateTime dt;
                //int cYY = global.cnfYear;
                //int cMM = global.cnfMonth;

                //if (t.日 < 16)
                //{
                //    cMM++;

                //    if (cMM > 12)
                //    {
                //        cYY++;
                //        cMM = 1;
                //    }
                //}

                //string cDt = cYY + "/" + cMM + "/" + t.日;

                DateTime dt = getCalenderDate(global.cnfYear, global.cnfMonth, t.日);

                if (dt == global.NODATE)
                {
                    dGV[cDay, mRow].Value = string.Empty;
                    dGV[cWeek, mRow].Value = string.Empty;

                    global.ChangeValueStatus = false;           // これ以下ChangeValueイベントを発生させない

                    dGV[cKinmuTaikei, mRow].Value = string.Empty;
                    dGV[cFlg, mRow].Value = global.flgOff;
                    dGV[cSH, mRow].Value = string.Empty;
                    dGV[cSM, mRow].Value = string.Empty;
                    dGV[cEH, mRow].Value = string.Empty;
                    dGV[cEM, mRow].Value = string.Empty;
                    dGV[cRH, mRow].Value = string.Empty;
                    dGV[cRM, mRow].Value = string.Empty;
                    dGV[cWH, mRow].Value = string.Empty;
                    dGV[cWM, mRow].Value = string.Empty;
                    dGV[cJ1, mRow].Value = string.Empty;
                    dGV[cJ2, mRow].Value = string.Empty;
                    dGV.Rows[mRow].ReadOnly = true;
                }
                else
                {
                    // 存在する日付と認識された場合、データを表示する
                    dGV[cDay, mRow].Value = t.日;
                    dGV[cWeek, mRow].Value = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);

                    global.ChangeValueStatus = false;           // これ以下ChangeValueイベントを発生させない

                    // 時刻区切り文字
                    dGV[cSE, mRow].Value = ":";
                    dGV[cEE, mRow].Value = ":";
                    dGV[cRE, mRow].Value = ":";
                    dGV[cWE, mRow].Value = ":";

                    dGV[cKinmuTaikei, mRow].Value = t.勤務体系コード;
                    dGV[cFlg, mRow].Value =  t.残業休日出勤申請;
                    dGV[cSH, mRow].Value = t.開始時;
                    dGV[cSM, mRow].Value = t.開始分;
                    dGV[cEH, mRow].Value = t.退勤時;
                    dGV[cEM, mRow].Value = t.退勤分;
                    dGV[cRH, mRow].Value = t.休憩時;
                    dGV[cRM, mRow].Value = t.休憩分;
                    dGV[cWH, mRow].Value = t.実労時;
                    dGV[cWM, mRow].Value = t.実労分;
                    dGV[cJ1, mRow].Value = t.事由１;
                    dGV[cJ2, mRow].Value = t.事由2;
                    dGV.Rows[mRow].ReadOnly = false;
                }

                dGV[cID, mRow].Value = t.ID.ToString();     // 明細ＩＤ
                global.ChangeValueStatus = true;            // ChangeValueStatusをtrueに戻す


                // 編集箇所をバックカラー表示
                foreach (var item in dts.出勤簿編集ログ.Where(a => a.勤務表ヘッダID == hID && a.行番号 == mRow + 1))
                {
                    if (item.Is列名Null())
                    {
                        continue;
                    }

                    if (global.cnfEditBackColor != "")
                    { 
                        dGV[item.列名, mRow].Style.BackColor = Color.FromArgb(Convert.ToInt32(global.cnfEditBackColor, 16));
                    }

                    //if (item.項目名 == CELL_CHECK)
                    //{
                    //    dGV[cFlg, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                    //else if (item.列名 == CELL_TAIKEICD)
                    //{
                    //    dGV[cKinmuTaikei, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                    //else if (item.列名 == CELL_KAISHI)
                    //{
                    //    dGV[cSH, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                    //else if (item.列名 == CELL_KAISHI_M)
                    //{
                    //    dGV[cSM, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                    //else if (item.列名 == CELL_TAIKIN)
                    //{
                    //    dGV[cEH, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                    //else if (item.列名 == CELL_TAIKIN_M)
                    //{
                    //    dGV[cEM, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                    //else if (item.列名 == CELL_REST)
                    //{
                    //    dGV[cRH, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                    //else if (item.列名 == CELL_REST_M)
                    //{
                    //    dGV[cRM, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                    //else if (item.列名 == CELL_WORK)
                    //{
                    //    dGV[cWH, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                    //else if (item.列名 == CELL_WORK_M)
                    //{
                    //    dGV[cWM, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                    //else if (item.列名 == CELL_JIYU1)
                    //{
                    //    dGV[cJ1, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                    //else if (item.列名 == CELL_JIYU2)
                    //{
                    //    dGV[cJ2, mRow].Style.BackColor = Color.Cornsilk;
                    //}
                }

                // 行インデックス加算
                mRow++;
            }

            //カレントセル選択状態としない
            dGV.CurrentCell = null;
        }
        
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     画像を表示する </summary>
        /// <param name="pic">
        ///     pictureBoxオブジェクト</param>
        /// <param name="imgName">
        ///     イメージファイルパス</param>
        /// <param name="fX">
        ///     X方向のスケールファクター</param>
        /// <param name="fY">
        ///     Y方向のスケールファクター</param>
        ///------------------------------------------------------------------------------------
        private void ImageGraphicsPaint(PictureBox pic, string imgName, float fX, float fY, int RectDest, int RectSrc)
        {
            Image _img = Image.FromFile(imgName);
            Graphics g = Graphics.FromImage(pic.Image);

            // 各変換設定値のリセット
            g.ResetTransform();

            // X軸とY軸の拡大率の設定
            g.ScaleTransform(fX, fY);

            // 画像を表示する
            g.DrawImage(_img, RectDest, RectSrc);

            // 現在の倍率,座標を保持する
            global.ZOOM_NOW = fX;
            global.RECTD_NOW = RectDest;
            global.RECTS_NOW = RectSrc;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     フォーム表示初期化 </summary>
        /// <param name="sID">
        ///     過去データ表示時のヘッダID</param>
        /// <param name="cIx">
        ///     勤務票ヘッダカレントレコードインデックス</param>
        ///------------------------------------------------------------------------------------
        private void formInitialize(string sID)
        {
            // テキストボックス表示色設定
            txtYear.BackColor = Color.White;   
            txtMonth.BackColor = Color.White;
            txtCode.BackColor = Color.White;
            chkKakunin.BackColor = SystemColors.Control;

            txtYear.ForeColor = global.defaultColor;
            txtMonth.ForeColor = global.defaultColor;
            txtCode.ForeColor = global.defaultColor;
            txtMemo.ForeColor = global.defaultColor;

            // ヘッダ情報表示欄
            txtYear.Text = string.Empty;
            txtMonth.Text = string.Empty;
            txtCode.Text = string.Empty;
            lblNoImage.Visible = false;

            // ヘッダ情報
            txtYear.ReadOnly = true;
            txtMonth.ReadOnly = true;
            txtCode.ReadOnly = true;

            // スクロールバー設定
            hScrollBar1.Enabled = true;
            hScrollBar1.Minimum = 0;
            hScrollBar1.Maximum = 0;
            hScrollBar1.Value = 0;
            hScrollBar1.LargeChange = 1;
            hScrollBar1.SmallChange = 1;

            //移動ボタン制御
            btnFirst.Enabled = false;
            btnNext.Enabled = false;
            btnBefore.Enabled = false;
            btnEnd.Enabled = false;

            //データ数表示
            //lblPage.Text = string.Empty;

            chkKakunin.Enabled = false;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     エラー表示 </summary>
        /// <param name="ocr">
        ///     OCRDATAクラス</param>
        ///------------------------------------------------------------------------------------
        private void ErrShow(OCRData ocr)
        {
            if (ocr._errNumber != ocr.eNothing)
            {
                // グリッドビューCellEnterイベント処理は実行しない
                gridViewCellEnterStatus = false;

                lblErrMsg.Visible = true;
                lblErrMsg.Text = ocr._errMsg;

                // 確認チェック
                if (ocr._errNumber == ocr.eDataCheck)
                {
                    chkKakunin.BackColor = Color.Yellow;
                    chkKakunin.Focus();
                }

                // 対象年月
                if (ocr._errNumber == ocr.eYearMonth)
                {
                    txtYear.BackColor = Color.Yellow;
                    txtMonth.BackColor = Color.Yellow;
                    txtYear.Focus();
                }

                // 対象月
                if (ocr._errNumber == ocr.eMonth)
                {
                    txtMonth.BackColor = Color.Yellow;
                    txtMonth.Focus();
                }

                // 社員番号
                if (ocr._errNumber == ocr.eShainNo)
                {
                    txtCode.BackColor = Color.Yellow;
                    txtCode.Focus();
                }

                // 対象日
                if (ocr._errNumber == ocr.eDay)
                {
                    //txtDay.BackColor = Color.Yellow;
                    //txtDay.Focus();
                }

                // 残業申請書チェック
                if (ocr._errNumber == ocr.eCFlg)
                {
                    dGV[cFlg, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cFlg, ocr._errRow];
                }

                // 勤務体系コード
                if (ocr._errNumber == ocr.eKinmuTaikeiCode)
                {
                    dGV[cKinmuTaikei, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cKinmuTaikei, ocr._errRow];
                }

                // 開始時
                if (ocr._errNumber == ocr.eSH)
                {
                    dGV[cSH, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cSH, ocr._errRow];
                }

                // 開始分
                if (ocr._errNumber == ocr.eSM)
                {
                    dGV[cSM, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cSM, ocr._errRow];
                }

                // 終了時
                if (ocr._errNumber == ocr.eEH)
                {
                    dGV[cEH, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cEH, ocr._errRow];
                }

                // 終了分
                if (ocr._errNumber == ocr.eEM)
                {
                    dGV[cEM, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cEM, ocr._errRow];
                }

                // 休憩・時
                if (ocr._errNumber == ocr.eRh)
                {
                    dGV[cRH, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cRH, ocr._errRow];
                }

                // 休憩・分
                if (ocr._errNumber == ocr.eRm)
                {
                    dGV[cRM, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cRM, ocr._errRow];
                }

                // 実労・時
                if (ocr._errNumber == ocr.eWh)
                {
                    dGV[cWH, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cWH, ocr._errRow];
                }

                // 実労・分
                if (ocr._errNumber == ocr.eWm)
                {
                    dGV[cWM, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cWM, ocr._errRow];
                }

                // 事由１
                if (ocr._errNumber == ocr.eJiyu1)
                {
                    dGV[cJ1, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cJ1, ocr._errRow];
                }

                // 事由2
                if (ocr._errNumber == ocr.eJiyu2)
                {
                    dGV[cJ2, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cJ2, ocr._errRow];
                }

                // グリッドビューCellEnterイベントステータスを戻す
                gridViewCellEnterStatus = true;

            }
        }
    }
}
