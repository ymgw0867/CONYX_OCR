using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CONYX_OCR.common;

namespace CONYX_OCR.config
{
    public class getConfig
    {
        CONYXDataSetTableAdapters.環境設定TableAdapter adp = new CONYXDataSetTableAdapters.環境設定TableAdapter();
        CONYXDataSet.環境設定DataTable cTbl = new CONYXDataSet.環境設定DataTable(); 

        public getConfig()
        {
            try
            {
                adp.Fill(cTbl);
                CONYXDataSet.環境設定Row r = cTbl.FindByID(global.configKEY);

                global.cnfYear = r.年;
                global.cnfMonth = r.月;
                global.cnfPath = r.汎用データ出力先;
                global.cnfArchived = r.データ保存月数;
                global.cnfMarume = r.丸め単位;

                if (r.Is編集済み背景色Null())
                {
                    global.cnfEditBackColor = "";
                }
                else
                {
                    global.cnfEditBackColor = r.編集済み背景色;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定年月取得", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
            }
        }
    }
}
