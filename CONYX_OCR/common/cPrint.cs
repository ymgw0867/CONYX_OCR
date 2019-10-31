using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Printing;
using Leadtools.WinForms;

namespace CONYX_OCR.common
{
    class cPrint
    {
        //プリント制御
        private int PRINT_Den;                          //印刷伝票
        private int PRINTMAXGYOU = 5;                   //最大印刷伝票数
        private int PRINTFONTSIZE = 8;                  //印刷フォントサイズ
        private int PrintMode;                          //全部印刷、一枚印刷の区分
        private int PrintPage = 1;                      //頁カウント
        private int Loopcnt = 0;                        //印刷データ数カウント
        private int wrkFirstDen = 0;
        private decimal KariSum = 0;
        private decimal KashiSum = 0;

        private float PrnX;                             //印刷位置X
        private float PrnY;                             //印刷位置Y
        private RasterImageViewer prnImage;             //印刷するLeadTools画像


        /// --------------------------------------------------------------------
        /// <summary>
        ///     伝票画像を印刷する </summary>
        /// <param name="Img">
        ///     LeadTools画像</param>
        /// --------------------------------------------------------------------
        public void Image(RasterImageViewer Img)
        {
            prnImage = Img;
            PrintDocument PrnImg = new PrintDocument();
            PrnImg.PrinterSettings = new PrinterSettings();

            //用紙方向：ヨコ
            PrnImg.DefaultPageSettings.Landscape = false;

            //用紙サイズ：A4
            foreach (System.Drawing.Printing.PaperSize ps in PrnImg.PrinterSettings.PaperSizes)
            {
                if (ps.Kind == System.Drawing.Printing.PaperKind.A4)
                {
                    PrnImg.DefaultPageSettings.PaperSize = ps;
                    break;
                }
            }

            //印刷実行
            PrnImg.PrintPage += new PrintPageEventHandler(Image_PrintPage);
            PrnImg.Print();
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     伝票画像印刷イベントハンドラ </summary>
        /// <param name="sender">
        ///     </param>
        /// <param name="e">
        ///     </param>
        /// --------------------------------------------------------------------
        private void Image_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //画像印刷
            int savePage = prnImage.Image.Page;

            try
            {
                using (Image img = prnImage.Image.ConvertToGdiPlusImage())
                {
                    e.Graphics.DrawImage(img, 0, 0);
                }
            }
            catch (Exception eX)
            {
                MessageBox.Show("伝票画像印刷中に不具合が発生したため印刷を中断します" + Environment.NewLine + eX.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                e.HasMorePages = false;
            }
            return;
        }

    }
}
