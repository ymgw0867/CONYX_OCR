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
    public partial class frmImgView : Form
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
        public frmImgView(string _sImg)
        {
            InitializeComponent();
            sImg = _sImg;
        }
        

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

        #region 終了ステータス定数
        const string END_BUTTON = "btn";
        const string END_MAKEDATA = "data";
        const string END_CONTOROL = "close";
        const string END_NODATA = "non Data";
        #endregion

        string dID = string.Empty;              // 表示する過去データのID
        string sImg = string.Empty;
        string sDBNM = string.Empty;            // データベース名
        

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);
            
            // 画像フォルダパス
            string dirName = Properties.Settings.Default.tifPath + sImg.Substring(0, 6) + @"\";

            // 画像表示 
            ShowImage(dirName + sImg);

            // tagを初期化
            this.Tag = string.Empty;
        }
        
        
        private void frmCorrect_Shown(object sender, EventArgs e)
        {
            if (dID != string.Empty) btnRtn.Focus();
        }
        

        private void btnNext_Click(object sender, EventArgs e)
        {

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
            // フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            ////「受入データ作成終了」「勤務票データなし」以外での終了のとき
            //if (this.Tag.ToString() != END_MAKEDATA && this.Tag.ToString() != END_NODATA)
            //{
            //    if (MessageBox.Show("終了します。よろしいですか", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            //    {
            //        e.Cancel = true;
            //        return;
            //    }
            //}
            
            // 解放する
            this.Dispose();
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
        
        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

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
