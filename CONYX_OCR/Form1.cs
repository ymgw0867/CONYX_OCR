using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CONYX_OCR.common;
using Leadtools;
using Leadtools.Codecs;
using Leadtools.ImageProcessing;
using Leadtools.ImageProcessing.Core;

namespace CONYX_OCR
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // 自分のコンピュータの登録がスキャン用ＰＣに登録されているか
            getPcName();

            // ＯＣＲ実施ＰＣか？
            if (Properties.Settings.Default.ocrStatus == global.flgOn)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }

            // 環境設定項目よみこみ
            config.getConfig cnf = new config.getConfig();
        }


        ///-------------------------------------------------------------------------------
        /// <summary>
        ///     自分のコンピュータの登録がスキャン用ＰＣに登録されているか調べる
        ///     登録済みのとき：「勤怠データ作成」「過去データ検索」「環境設定」のボタン True
        ///     未登録のとき：「勤怠データ作成」「過去データ検索」「環境設定」のボタン false
        /// </summary>
        ///-------------------------------------------------------------------------------
        private void getPcName()
        {
            string pcName = string.Empty;

            // 登録されていないとき終了します
            pcName = Utility.getPcDir();
            if (pcName == string.Empty)
            {
                MessageBox.Show("このコンピュータがＯＣＲ出力先として登録されていません。", "出力先未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.button2.Enabled = false;
                this.button3.Enabled = false;
                this.button5.Enabled = false;
            }
            else
            {
                this.button2.Enabled = true;
                this.button3.Enabled = true;
                this.button5.Enabled = true;
                global.pcName = pcName;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            config.frmMsOutPath frm = new config.frmMsOutPath();
            frm.ShowDialog();
            this.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Hide();

            // 出力先ＰＣ選択画面
            OCR.frmOCRPC frm = new OCR.frmOCRPC();
            frm.ShowDialog();
            string pcName = frm._outPC;
            frm.Dispose();

            if (pcName == string.Empty)
            {
                Show();
                return;
            }

            if (MessageBox.Show("出勤簿画像のＯＣＲ認識を行います。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                Show();
                return;
            }

            // ＯＣＲ認識実行
            Hide();
            doFaxOCR(Properties.Settings.Default.wrHands_Job, Properties.Settings.Default.dataPath);

            // PC毎の出力先フォルダがなければ作成する
            string rPath = Properties.Settings.Default.pcPath + pcName + @"\";
            if (System.IO.Directory.Exists(rPath) == false)
            {
                System.IO.Directory.CreateDirectory(rPath);
            }

            // データを移動する
            foreach (var file in System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath))
            {
                System.IO.File.Move(file, rPath + System.IO.Path.GetFileName(file));
            }

            Show();
        }


        private void doFaxOCR(string wrJobName, string outPath)
        {
            int cnt = System.IO.Directory.GetFiles(Properties.Settings.Default.scanPath, "*.tif").Count();
            if (cnt == 0)
            {
                MessageBox.Show("ＯＣＲ認識対象画像がありません", "OCR認識", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                // ファイル名（日付時間部分）
                string fName = string.Format("{0:0000}", DateTime.Today.Year) +
                        string.Format("{0:00}", DateTime.Today.Month) +
                        string.Format("{0:00}", DateTime.Today.Day) +
                        string.Format("{0:00}", DateTime.Now.Hour) +
                        string.Format("{0:00}", DateTime.Now.Minute) +
                        string.Format("{0:00}", DateTime.Now.Second);

                int dNum = 0;                       // ファイル名末尾連番

                /* マルチTiff画像をシングルtifに分解後にSCANフォルダ → TRAYフォルダ */
                if (MultiTif(Properties.Settings.Default.scanPath, Properties.Settings.Default.trayPath, fName))
                {
                    // WinReaderを起動して出勤簿をスキャンしてOCR処理を実施する
                    WinReaderOCR(wrJobName);

                    /* OCR認識結果ＣＳＶデータを出勤簿ごとに分割して
                     * 画像ファイルと共にDATAフォルダへ移動する */
                    LoadCsvDivide(fName, ref dNum, outPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
                MessageBox.Show("終了しました");
            }
        }

        ///------------------------------------------------------------------------------
        /// <summary>
        ///     マルチフレームの画像ファイルを頁ごとに分割する </summary>
        /// <param name="InPath">
        ///     画像ファイル入力パス</param>
        /// <param name="outPath">
        ///     分割後出力パス</param>
        /// <returns>
        ///     true:分割を実施, false:分割ファイルなし</returns>
        ///------------------------------------------------------------------------------
        private bool MultiTif(string InPath, string outPath, string fName)
        {
            //スキャン出力画像を確認
            if (System.IO.Directory.GetFiles(InPath, "*.tif").Count() == 0)
            {
                return false;
            }

            // 出力先フォルダがなければ作成する
            if (System.IO.Directory.Exists(outPath) == false)
            {
                System.IO.Directory.CreateDirectory(outPath);
            }

            // 出力先フォルダ内の全てのファイルを削除する（通常ファイルは存在しないが例外処理などで残ってしまった場合に備えて念のため）
            foreach (string files in System.IO.Directory.GetFiles(outPath, "*"))
            {
                System.IO.File.Delete(files);
            }

            RasterCodecs.Startup();
            RasterCodecs cs = new RasterCodecs();

            int _pageCount = 0;
            string fnm = string.Empty;

            //コマンドを準備します。(傾き・ノイズ除去・リサイズ)
            DeskewCommand Dcommand = new DeskewCommand();
            DespeckleCommand Dkcommand = new DespeckleCommand();
            SizeCommand Rcommand = new SizeCommand();

            // オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            int cImg = System.IO.Directory.GetFiles(InPath, "*.tif").Count();
            int cCnt = 0;

            // マルチTIFを分解して画像ファイルをTRAYフォルダへ保存する
            foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                cCnt++;

                //プログレスバー表示
                frmP.Text = "OCR変換画像データロード中　" + cCnt.ToString() + "/" + cImg;
                frmP.progressValue = cCnt * 100 / cImg;
                frmP.ProgressStep();

                // 画像読み出す
                RasterImage leadImg = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, -1);

                // 頁数を取得
                int _fd_count = leadImg.PageCount;

                // 頁ごとに読み出す
                for (int i = 1; i <= _fd_count; i++)
                {
                    //ページを移動する
                    leadImg.Page = i;

                    // ファイル名設定
                    _pageCount++;
                    fnm = outPath + fName + string.Format("{0:000}", _pageCount) + ".tif";

                    //画像補正処理　開始 ↓ ****************************
                    try
                    {
                        //画像の傾きを補正します。
                        Dcommand.Flags = DeskewCommandFlags.DeskewImage | DeskewCommandFlags.DoNotFillExposedArea;
                        Dcommand.Run(leadImg);
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(i + "画像の傾き補正エラー：" + ex.Message);
                    }

                    //ノイズ除去
                    try
                    {
                        Dkcommand.Run(leadImg);
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(i + "ノイズ除去エラー：" + ex.Message);
                    }

                    ////解像度調整(200*200dpi)
                    //leadImg.XResolution = 200;
                    //leadImg.YResolution = 200;

                    ////A4縦サイズに変換(ピクセル単位)
                    //Rcommand.Width = 1637;
                    //Rcommand.Height = 2322;
                    //try
                    //{
                    //    Rcommand.Run(leadImg);
                    //}
                    //catch (Exception ex)
                    //{
                    //    //MessageBox.Show(i + "解像度調整エラー：" + ex.Message);
                    //}

                    //画像補正処理　終了↑ ****************************

                    // 画像保存
                    cs.Save(leadImg, fnm, RasterImageFormat.Tif, 0, i, i, 1, CodecsSavePageMode.Insert);
                }
            }

            //LEADTOOLS入出力ライブラリを終了します。
            RasterCodecs.Shutdown();

            // InPathフォルダの全てのtifファイルを削除する
            foreach (var files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                System.IO.File.Delete(files);
            }

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            return true;
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     WinReaderを起動して出勤簿をスキャンしてOCR処理を実施する </summary>
        ///----------------------------------------------------------------------------------
        private void WinReaderOCR(string wrJobName)
        {
            // WinReaderJOB起動文字列
            string JobName = @"""" + wrJobName + @"""" + " /H2";
            string winReader_exe = Properties.Settings.Default.wrHands_Path +
                Properties.Settings.Default.wrHands_Prg;

            // ProcessStartInfo の新しいインスタンスを生成する
            System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo();

            // 起動するアプリケーションを設定する
            p.FileName = winReader_exe;

            // コマンドライン引数を設定する（WinReaderのJOB起動パラメーター）
            p.Arguments = JobName;

            // WinReaderを起動します
            System.Diagnostics.Process hProcess = System.Diagnostics.Process.Start(p);

            // taskが終了するまで待機する
            hProcess.WaitForExit();
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     伝票ＣＳＶデータを一枚ごとに分割する </summary>
        ///-----------------------------------------------------------------
        private void LoadCsvDivide(string fnm, ref int dNum, string outPath)
        {
            string imgName = string.Empty;      // 画像ファイル名
            string firstFlg = global.FLGON;
            string[] stArrayData;               // CSVファイルを１行単位で格納する配列
            string newFnm = string.Empty;
            int dCnt = 0;   // 処理件数

            // 対象ファイルの存在を確認します
            if (!System.IO.File.Exists(Properties.Settings.Default.readPath + Properties.Settings.Default.wrReaderOutFile))
            {
                return;
            }

            // StreamReader の新しいインスタンスを生成する
            //入力ファイル
            System.IO.StreamReader inFile = new System.IO.StreamReader(Properties.Settings.Default.readPath + Properties.Settings.Default.wrReaderOutFile, Encoding.Default);

            // 読み込んだ結果をすべて格納するための変数を宣言する
            string stResult = string.Empty;
            string stBuffer;

            // 行番号
            int sRow = 0;

            // 読み込みできる文字がなくなるまで繰り返す
            while (inFile.Peek() >= 0)
            {
                // ファイルを 1 行ずつ読み込む
                stBuffer = inFile.ReadLine();

                // カンマ区切りで分割して配列に格納する
                stArrayData = stBuffer.Split(',');

                //先頭に「*」があったら新たな伝票なのでCSVファイル作成
                if ((stArrayData[0] == "*"))
                {
                    //最初の伝票以外のとき
                    if (firstFlg != global.FLGON)
                    {
                        //ファイル書き出し
                        outFileWrite(stResult, Properties.Settings.Default.readPath + imgName, outPath + newFnm);
                    }

                    firstFlg = global.FLGOFF;

                    // 伝票連番
                    dNum++;

                    // 処理件数
                    dCnt++;

                    // ファイル名
                    newFnm = fnm + dNum.ToString().PadLeft(3, '0');

                    //画像ファイル名を取得
                    imgName = stArrayData[1];

                    //文字列バッファをクリア
                    stResult = string.Empty;

                    // 文字列再校正（画像ファイル名を変更する）
                    stBuffer = string.Empty;
                    for (int i = 0; i < stArrayData.Length; i++)
                    {
                        if (stBuffer != string.Empty)
                        {
                            stBuffer += ",";
                        }

                        // 画像ファイル名を変更する
                        if (i == 1)
                        {
                            stArrayData[i] = newFnm + ".tif"; // 画像ファイル名を変更
                        }

                        //// 日付（６桁）を年月日（２桁毎）に分割する
                        //if (i == 3)
                        //{
                        //    string dt = stArrayData[i].PadLeft(6, '0');
                        //    stArrayData[i] = dt.Substring(0, 2) + "," + dt.Substring(2, 2) + "," + dt.Substring(4, 2);
                        //}

                        // フィールド結合
                        stBuffer += stArrayData[i];
                    }

                    sRow = 0;
                }
                else
                {
                    sRow++;
                }

                // 読み込んだものを追加で格納する
                stResult += (stBuffer + Environment.NewLine);

                //// 最終行は追加しない（伝票区別記号(*)のため）
                //if (sRow <= global.MAXGYOU_PRN)
                //{
                //    // 読み込んだものを追加で格納する
                //    stResult += (stBuffer + Environment.NewLine);
                //}
            }

            // 後処理
            if (dNum > 0)
            {
                //ファイル書き出し
                outFileWrite(stResult, Properties.Settings.Default.readPath + imgName, outPath + newFnm);

                // 入力ファイルを閉じる
                inFile.Close();

                //入力ファイル削除 : "txtout.csv"
                Utility.FileDelete(Properties.Settings.Default.readPath, Properties.Settings.Default.wrReaderOutFile);

                //画像ファイル削除 : "WRH***.tif"
                Utility.FileDelete(Properties.Settings.Default.readPath, "WRH*.tif");
            }
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     分割ファイルを書き出す </summary>
        /// <param name="tempResult">
        ///     書き出す文字列</param>
        /// <param name="tempImgName">
        ///     元画像ファイルパス</param>
        /// <param name="outFileName">
        ///     新ファイル名</param>
        ///----------------------------------------------------------------------------
        private void outFileWrite(string tempResult, string tempImgName, string outFileName)
        {
            //出力ファイル
            //System.IO.StreamWriter outFile = new System.IO.StreamWriter(Properties.Settings.Default.dataPath + outFileName + ".csv",
            //                                        false, System.Text.Encoding.GetEncoding(932));

            // 2017/11/20
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(outFileName + ".csv", false, System.Text.Encoding.GetEncoding(932));

            // ファイル書き出し
            outFile.Write(tempResult);

            //ファイルクローズ
            outFile.Close();

            //画像ファイルをコピー
            //System.IO.File.Copy(tempImgName, Properties.Settings.Default.dataPath + outFileName + ".tif");

            // 2017/11/20
            System.IO.File.Copy(tempImgName, outFileName + ".tif");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Hide();
            config.frmConfig frm = new config.frmConfig();
            frm.ShowDialog();
            Show();

            // 環境設定項目よみこみ
            config.getConfig cnf = new config.getConfig();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 環境設定年月の確認
            string msg = "処理対象年月は " + global.cnfYear.ToString() + "年 " + global.cnfMonth.ToString() + "月です。よろしいですか？";
            if (MessageBox.Show(msg, "勤務データ登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;

            this.Hide();
            frmComSelect frm = new frmComSelect();
            frm.ShowDialog();

            if (frm._pblDbName != string.Empty)
            {
                // 選択領域のデータベース名を取得します
                string _ComName = frm._pblComName;
                string _ComDBName = frm._pblDbName;
                frm.Dispose();

                // 出勤簿データ作成画面
                OCR.frmCorrect frmg = new OCR.frmCorrect(_ComDBName, _ComName, string.Empty);
                frmg.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Hide();
            OCR.frmUnSubmit frm = new OCR.frmUnSubmit();
            frm.ShowDialog();
            Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Hide();
            OCR.frmEditLogRep frm = new OCR.frmEditLogRep();
            frm.ShowDialog();
            Show();                
        }
    }
}
