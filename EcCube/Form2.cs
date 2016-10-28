using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace EcCube
{
    //--------------------------------------------------------------
    #region 宣言
    /// <summary>
    /// 表示項目要素
    /// </summary>
    public enum EnumShowItem
    {
        /// <summary>有効・無効</summary>
        ENABLE = 0,
        /// <summary>注文番号</summary>
        ORDER_ID,
        /// <summary>購入者名</summary>
        NAME,
        /// <summary>製品名リスト</summary>
        PRODUCT_NAME_LIST,
        /// <summary>箱数</summary>
        BOX_NUM,

        MAX
    }
    #endregion
    //--------------------------------------------------------------
    public partial class Form2 : Form
    {
        //--------------------------------------------------------------
        // 読み込みデーター
        List<string[]> m_listReadData = new List<string[]>();
        // 注文データーリスト
        List<CEcCubeData> m_listOrderData = new List<CEcCubeData>();
        // ゆうプリR出力ファイル名
        string m_YpprFileName;

        /// <summary>グリッド列の幅</summary>
        int[] m_iGridColWidth = new int[(int)EnumShowItem.MAX];
        /// <summary>表示項目要素文字列定数</summary>
        string[] SETTING_SHOW_ITEM_STR = new string[(int)EnumShowItem.MAX];

        string m_ListProductFileName;
        List<CProductData> m_listProduct = new List<CProductData>();
        //--------------------------------------------------------------
        public Form2()
        {
            InitializeComponent();

            m_iGridColWidth[(int)EnumShowItem.ENABLE] = 50;
            m_iGridColWidth[(int)EnumShowItem.ORDER_ID] = 170;
            m_iGridColWidth[(int)EnumShowItem.NAME] = 140;
            m_iGridColWidth[(int)EnumShowItem.PRODUCT_NAME_LIST] = 320;
            m_iGridColWidth[(int)EnumShowItem.BOX_NUM] = 50;

            SETTING_SHOW_ITEM_STR[(int)EnumShowItem.ENABLE] = "有効";
            SETTING_SHOW_ITEM_STR[(int)EnumShowItem.ORDER_ID] = "注文番号";
            SETTING_SHOW_ITEM_STR[(int)EnumShowItem.NAME] = "購入者名";
            SETTING_SHOW_ITEM_STR[(int)EnumShowItem.PRODUCT_NAME_LIST] = "製品名";
            SETTING_SHOW_ITEM_STR[(int)EnumShowItem.BOX_NUM] = "箱数";
        }
        //--------------------------------------------------------------
        public void ShowDlg(string strFileName)
        {
            if (LoadDataCsv(strFileName) == false)
            {
                MessageBox.Show("読み込みエラー");
                return;
            }
            // 注文データー作成
            MakeOrderData();

            // ゆうプリRファイル名
            m_YpprFileName = Path.GetDirectoryName(strFileName) + "\\" + "ゆうプリR.csv";

            m_ListProductFileName = Path.GetDirectoryName(strFileName) + "\\" + "今日の出荷.csv";


            // 画面表示
            ShowDialog();
        }
        //--------------------------------------------------------------
        // ファイルからデータ読み込み
        public bool LoadDataCsv(string strFileName)
        {
            // ヘッダー
            string[] strHeader;

            // ファイルを読み込む
            StreamReader clsSr;
            try
            {
                clsSr = new StreamReader(strFileName, System.Text.Encoding.GetEncoding("shift_jis"));
                try
                {
                    string strLine;

                    // 1行目読み込み
                    if ((strLine = clsSr.ReadLine()) == null)
                    {
                        throw new FormatException();
                    }
                    // ヘッダーの文字列を読み、種類を判断する
                    strHeader = strLine.Split(',');
                    if (strHeader[0] != "注文番号")
                    {
                        throw new FormatException();
                    }
                    // データーを読み込む                
                    while ((strLine = clsSr.ReadLine()) != null)
                    {
                        if (strLine == "")
                        {
                            break;
                        }
                        // カンマが含まれている場合は取り除く
                        bool bCheck = false;
                        for (int i = 0; i < strLine.Length; i++)
                        {
                            if (bCheck == false)
                            {
                                if (strLine[i] == '\"')
                                {
                                    bCheck = true;
                                }
                            }
                            else
                            {
                                if (strLine[i] == ',')
                                {
                                    strLine = strLine.Substring(0, i) + " " + strLine.Substring(i + 1, strLine.Length - (i + 1));
                                }
                                else if (strLine[i] == '\"')
                                {
                                    bCheck = false;
                                }
                            }
                        }
                        // ダブルコーテーションが含まれている場合は取り除く
                        strLine = strLine.Replace("\"", "");

                        // 全てのデーターを格納
                        m_listReadData.Add(strLine.Split(','));
                    }
                }
                catch
                {
                    return false;
                }
                finally
                {
                    clsSr.Close();
                }
            }
            catch
            {
                return false;
            }

            return true;
        }
        //--------------------------------------------------------------
        void MakeOrderData()
        {
            // 注文データーリストに追加する
            int iIDPos = (int)EnumEcCubeItem.ORDER_ID;

            m_listOrderData.Clear();

            while (m_listReadData.Count > 0)
            {
                CEcCubeData clsEcCubeData = new CEcCubeData();

                clsEcCubeData.strOrderId = m_listReadData[0][(int)EnumEcCubeItem.ORDER_ID];
                clsEcCubeData.strName = m_listReadData[0][(int)EnumEcCubeItem.FIRST_NAME] + m_listReadData[0][(int)EnumEcCubeItem.LAST_NAME];
                clsEcCubeData.strPostNo = m_listReadData[0][(int)EnumEcCubeItem.POST_NO_1] + "-" + m_listReadData[0][(int)EnumEcCubeItem.POST_NO_2];
                clsEcCubeData.strAddress1 = m_listReadData[0][(int)EnumEcCubeItem.ADDRESS_1];
                clsEcCubeData.strAddress2 = m_listReadData[0][(int)EnumEcCubeItem.ADDRESS_2];
                clsEcCubeData.strAddress3 = m_listReadData[0][(int)EnumEcCubeItem.ADDRESS_3];
                clsEcCubeData.strCompanyName = m_listReadData[0][(int)EnumEcCubeItem.COMPANY_NAME];
                clsEcCubeData.strPhoneNo = m_listReadData[0][(int)EnumEcCubeItem.PHONE_NO_1] + "-" + m_listReadData[0][(int)EnumEcCubeItem.PHONE_NO_2] + "-" + m_listReadData[0][(int)EnumEcCubeItem.PHONE_NO_3];
                clsEcCubeData.strDeliberyDate = m_listReadData[0][(int)EnumEcCubeItem.DELIBERY_DATE];
                clsEcCubeData.strDeliberyTime = m_listReadData[0][(int)EnumEcCubeItem.DELIBERY_TIME];
                clsEcCubeData.strTotalPay = m_listReadData[0][(int)EnumEcCubeItem.TOTAL_PAY];
                clsEcCubeData.strPaymentMethod = m_listReadData[0][(int)EnumEcCubeItem.PAYMENT_METHOD];
                clsEcCubeData.strDeliveryMethod = m_listReadData[0][(int)EnumEcCubeItem.DELIVERY_METHOD];
                clsEcCubeData.strEMail = m_listReadData[0][(int)EnumEcCubeItem.E_MAIL];
                clsEcCubeData.listProductCode.Add(m_listReadData[0][(int)EnumEcCubeItem.PRODUCT_CODE]);
                clsEcCubeData.listQuantity.Add(int.Parse(m_listReadData[0][(int)EnumEcCubeItem.QUANTITY]));
                clsEcCubeData.listProductName.Add(m_listReadData[0][(int)EnumEcCubeItem.PRODUCT_NAME]);

                // データ削除
                string strID = m_listReadData[0][iIDPos];
                for (int i = m_listReadData.Count - 1; i >= 0; i--)
                {
                    if (strID == m_listReadData[i][iIDPos])
                    {
                        clsEcCubeData.listProductCode.Add(m_listReadData[i][(int)EnumEcCubeItem.PRODUCT_CODE]);
                        clsEcCubeData.listQuantity.Add(int.Parse(m_listReadData[i][(int)EnumEcCubeItem.QUANTITY]));
                        clsEcCubeData.listProductName.Add(m_listReadData[i][(int)EnumEcCubeItem.PRODUCT_NAME]);
                        m_listReadData.RemoveAt(i);
                    }
                }
                // 最後の製品名リストは重複しているので削除する
                clsEcCubeData.listProductCode.RemoveAt(clsEcCubeData.listProductCode.Count - 1);
                clsEcCubeData.listQuantity.RemoveAt(clsEcCubeData.listQuantity.Count - 1);
                clsEcCubeData.listProductName.RemoveAt(clsEcCubeData.listProductName.Count - 1);

                // 箱数設定
                bool bSmallBoxFlag = false;
                clsEcCubeData.iTotalBoxNum = 0;
                for (int i = 0; i < clsEcCubeData.listProductCode.Count; i++)
                {
                    if (clsEcCubeData.listProductCode[i] == "mcn004")
                    {
                        // CO2本体
                        clsEcCubeData.iTotalBoxNum += (4 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn005")
                    {
                        // 水冷
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn014")
                    {
                        // 1.6W Mini
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn015")
                    {
                        // 3.5W Mini
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn016")
                    {
                        // 安全カバー
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn019")
                    {
                        // 拡張フレームセット
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn021")
                    {
                        // FABOOL Laser Mini用排気ファンキット
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else
                    {
                        // 小物は一つの箱にまとめる
                        bSmallBoxFlag = true;
                    }
                }
                if (bSmallBoxFlag == true)
                {
                    clsEcCubeData.iTotalBoxNum++;
                }
                if (clsEcCubeData.iTotalBoxNum == 0)
                {
                    clsEcCubeData.iTotalBoxNum = 1;
                }

                // 送り状種別設定
                if (clsEcCubeData.strPaymentMethod == "代金引換")
                {
                    clsEcCubeData.enInvoiceClass = EnumInvoiceClass.CASH_ON_DELIVERY;
                }
                else
                {
                    clsEcCubeData.enInvoiceClass = EnumInvoiceClass.DELIVERY;
                }

                // データ追加
                m_listOrderData.Add(clsEcCubeData);
            }
        }
        //--------------------------------------------------------------
        private void Form2_Load(object sender, EventArgs e)
        {
            #region データグリッドの初期設定
            // 左端の三角のプロパティ
            Dgv_Data.RowHeadersVisible = false;
            // ユーザーが新しい行を追加できないようにする
            Dgv_Data.AllowUserToAddRows = false;
            // ユーザーが削除できないようにする
            Dgv_Data.AllowUserToDeleteRows = false;
            // 列の幅は自動調整しない
            Dgv_Data.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            // セル内で文字列を折り返す
            Dgv_Data.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            // ヘッダーとすべてのセルの内容に合わせて、行の高さを自動調整する
            Dgv_Data.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            // 列の幅をユーザーが変更できないようにする
            Dgv_Data.AllowUserToResizeColumns = false;
            // 行の高さをユーザーが変更できないようにする
            Dgv_Data.AllowUserToResizeRows = false;
            // セル、行、列が複数選択されないようにする
            Dgv_Data.MultiSelect = false;
            //全ての列の背景色を変更する
            Dgv_Data.RowsDefaultCellStyle.BackColor = System.Drawing.Color.AntiqueWhite;
            //奇数行の色を変更する
            Dgv_Data.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.White;
            #endregion
            #region データグリッドの列設定
            // 列の設定
            Dgv_Data.Columns.Clear();
            Dgv_Data.Columns.Add(new DataGridViewCheckBoxColumn() { HeaderText = SETTING_SHOW_ITEM_STR[(int)EnumShowItem.ENABLE] });
            Dgv_Data.Columns.Add(new DataGridViewTextBoxColumn() { HeaderText = SETTING_SHOW_ITEM_STR[(int)EnumShowItem.ORDER_ID] });
            Dgv_Data.Columns.Add(new DataGridViewTextBoxColumn() { HeaderText = SETTING_SHOW_ITEM_STR[(int)EnumShowItem.NAME] });
            Dgv_Data.Columns.Add(new DataGridViewTextBoxColumn() { HeaderText = SETTING_SHOW_ITEM_STR[(int)EnumShowItem.PRODUCT_NAME_LIST] });
            Dgv_Data.Columns.Add(new DataGridViewTextBoxColumn() { HeaderText = SETTING_SHOW_ITEM_STR[(int)EnumShowItem.BOX_NUM] });

            for (int i = 0; i < (int)EnumShowItem.MAX; i++)
            {
                // ソートを全て無効
                Dgv_Data.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                // 文字列を左に表示
                Dgv_Data.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                Dgv_Data.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                // 幅を設定
                Dgv_Data.Columns[i].Width = m_iGridColWidth[i];
            }
            // 列の読み取り専用を設定
            Dgv_Data.Columns[(int)EnumShowItem.ORDER_ID].ReadOnly = true;
            Dgv_Data.Columns[(int)EnumShowItem.NAME].ReadOnly = true;
            Dgv_Data.Columns[(int)EnumShowItem.PRODUCT_NAME_LIST].ReadOnly = true;
            Dgv_Data.Columns[(int)EnumShowItem.BOX_NUM].ReadOnly = true;
            #endregion

            // データを全てクリア
            Dgv_Data.RowCount = 0;
            // 行数設定
            Dgv_Data.RowCount = m_listOrderData.Count;

            for (int i = 0; i < m_listOrderData.Count; i++)
            {
                // 有効・無効の設定
                Dgv_Data.Rows[i].Cells[(int)EnumShowItem.ENABLE].Value = m_listOrderData[i].bEnable;
                // 注文番号の設定
                Dgv_Data.Rows[i].Cells[(int)EnumShowItem.ORDER_ID].Value = m_listOrderData[i].strOrderId;
                // 購入者名の設定
                Dgv_Data.Rows[i].Cells[(int)EnumShowItem.NAME].Value = m_listOrderData[i].strName;
                // 製品名リストの設定
                for (int j = 0; j < m_listOrderData[i].listProductName.Count; j++)
                {
                    string strProductName = m_listOrderData[i].listProductName[j];
                    if (strProductName.Length > 25)
                    {
                        strProductName = strProductName.Substring(0, 25);
                    }

                    Dgv_Data.Rows[i].Cells[(int)EnumShowItem.PRODUCT_NAME_LIST].Value += strProductName;
                    if (j != m_listOrderData[i].listProductName.Count - 1)
                    {
                        Dgv_Data.Rows[i].Cells[(int)EnumShowItem.PRODUCT_NAME_LIST].Value += "\n";
                    }
                }

                // 箱数の設定
                Dgv_Data.Rows[i].Cells[(int)EnumShowItem.BOX_NUM].Value = m_listOrderData[i].iTotalBoxNum;
            }

        }
        //--------------------------------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            // データの取得
            for (int i = 0; i < m_listOrderData.Count; i++)
            {
                // 有効・無効の取得
                m_listOrderData[i].bEnable = (bool)Dgv_Data.Rows[i].Cells[(int)EnumShowItem.ENABLE].Value;
            }

            // コンバート処理
            if (ConvertCsv() == false)
            {
                MessageBox.Show("書き込みエラー");
                return;
            }

            // 商品出荷個数
            MakeProductData();
            if (WriteProductData() == false)
            {
                MessageBox.Show("書き込みエラー");
                return;
            }

            Close();
        }
        //--------------------------------------------------------------
        // コンバート処理実行
        public bool ConvertCsv()
        {
            // 宅急便・代引き
            if (ConvertDelivery() == false)
            {
                return true;
            }

            return true;
        }
        //--------------------------------------------------------------
        public bool ConvertDelivery()
        {
            for (int i = 0; i < m_listOrderData.Count; i++)
            {
                if (m_listOrderData[i].bEnable == false)
                {
                    continue;
                }

                // ゆうプリR
                string[] strYpprData = new string[(int)EnumYpprItem.MAX];

                // お客様側管理番号
                strYpprData[(int)EnumYpprItem.USER_ID] = "";
                // 発送予定日
                strYpprData[(int)EnumYpprItem.SHIPPING_SCHEDULE_DATE] = DateTime.Now.ToString("yyyyMMdd");
                // 発送予定時間
                strYpprData[(int)EnumYpprItem.SHIPPING_SCHEDULE_TIME] = "02";
                // 郵便種別
                strYpprData[(int)EnumYpprItem.POST_CLASS] = "0";
                // 支払元
                if (m_listOrderData[i].strPaymentMethod == "代金引換")
                {
                    strYpprData[(int)EnumYpprItem.PAYMENT_SOURCE] = "2";
                }
                else
                {
                    strYpprData[(int)EnumYpprItem.PAYMENT_SOURCE] = "0";
                }
                // 送り状種別
                if (m_listOrderData[i].strPaymentMethod == "代金引換")
                {
                    strYpprData[(int)EnumYpprItem.INVOICE_CLASS] = "1100783003";
                }
                else
                {
                    strYpprData[(int)EnumYpprItem.INVOICE_CLASS] = "1100783001";
                }
                // お届け先郵便番号
                strYpprData[(int)EnumYpprItem.TRANSPORT_POST_NO] = m_listOrderData[i].strPostNo;
                // お届け先住所
                strYpprData[(int)EnumYpprItem.TRANSPORT_ADDRESS_1] = m_listOrderData[i].strAddress1;
                strYpprData[(int)EnumYpprItem.TRANSPORT_ADDRESS_2] = m_listOrderData[i].strAddress2;
                strYpprData[(int)EnumYpprItem.TRANSPORT_ADDRESS_3] = m_listOrderData[i].strAddress3;

                if (m_listOrderData[i].strCompanyName == "")
                {
                    // お届け先名称1
                    strYpprData[(int)EnumYpprItem.TRANSPORT_NAME_1] = m_listOrderData[i].strName;
                    // お届け先名称2
                    strYpprData[(int)EnumYpprItem.TRANSPORT_NAME_2] = "";
                }
                else
                {
                    // お届け先名称1
                    strYpprData[(int)EnumYpprItem.TRANSPORT_NAME_1] = m_listOrderData[i].strCompanyName;
                    // お届け先名称2
                    strYpprData[(int)EnumYpprItem.TRANSPORT_NAME_2] = m_listOrderData[i].strName;
                }
                // お届け先敬称
                strYpprData[(int)EnumYpprItem.TRANSPORT_TITLE] = "0";
                // お届け先電話番号
                strYpprData[(int)EnumYpprItem.TRANSPORT_TEL] = m_listOrderData[i].strPhoneNo;
                // お届け先メール
                strYpprData[(int)EnumYpprItem.TRANSPORT_MAIL] = m_listOrderData[i].strEMail;
                // 発送元郵便番号
                strYpprData[(int)EnumYpprItem.ORIGIN_POST_NO] = "4000306";
                // 発送元住所1
                strYpprData[(int)EnumYpprItem.ORIGIN_ADDRESS_1] = "山梨県";
                // 発送元住所2
                strYpprData[(int)EnumYpprItem.ORIGIN_ADDRESS_2] = "南アルプス市小笠原";
                // 発送元住所3
                strYpprData[(int)EnumYpprItem.ORIGIN_ADDRESS_3] = "1589-1";
                // 発送元名称1
                strYpprData[(int)EnumYpprItem.ORIGIN_NAME_1] = "smartDIYs";
                // 発送元名称2
                strYpprData[(int)EnumYpprItem.ORIGIN_NAME_2] = "";
                // 発送元敬称
                strYpprData[(int)EnumYpprItem.ORIGIN_TITLE] = "0";
                // 発送元電話番号
                strYpprData[(int)EnumYpprItem.ORIGIN_TEL] = "05037867989";
                // 発送元メール
                strYpprData[(int)EnumYpprItem.ORIGIN_MAIL] = "support@smartdiys.com";
                // こわれもの
                strYpprData[(int)EnumYpprItem.BREAKABLE_FLG] = "1";
                // 逆さま厳禁
                strYpprData[(int)EnumYpprItem.WAY_UP_FLG] = "1";
                // 下積み厳禁
                strYpprData[(int)EnumYpprItem.DO_NOT_STACK_FLG] = "1";
                // 厚さ
                strYpprData[(int)EnumYpprItem.THICKNESS] = "";
                // お届け日
                if (m_listOrderData[i].strDeliberyDate != "")
                {
                    DateTime timeDelibery;
                    DateTime timeNow;

                    timeDelibery = new DateTime(int.Parse(m_listOrderData[i].strDeliberyDate.Substring(0, 4)),
                                                int.Parse(m_listOrderData[i].strDeliberyDate.Substring(5, 2)),
                                                int.Parse(m_listOrderData[i].strDeliberyDate.Substring(8, 2)),
                                                0, 0, 0
                                                );
                    timeNow = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);

                    if (timeDelibery.CompareTo(timeNow) > 0)
                    {
                        strYpprData[(int)EnumYpprItem.DELIBERY_DATE] = m_listOrderData[i].strDeliberyDate.Replace("/", "");
                    }
                    else
                    {
                        strYpprData[(int)EnumYpprItem.DELIBERY_DATE] = "";
                    }
                }
                else
                {
                    strYpprData[(int)EnumYpprItem.DELIBERY_DATE] = "";
                }
                // お届け時間
                if (m_listOrderData[i].strDeliberyTime == "午前中")
                {
                    strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "51";
                }
                else if (m_listOrderData[i].strDeliberyTime == "12～14")
                {
                    strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "52";
                }
                else if (m_listOrderData[i].strDeliberyTime == "14～16")
                {
                    strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "53";
                }
                else if (m_listOrderData[i].strDeliberyTime == "16～18")
                {
                    strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "54";
                }
                else if (m_listOrderData[i].strDeliberyTime == "18～20")
                {
                    strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "55";
                }
                else
                {
                    strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "00";
                }
                // 発行枚数
                strYpprData[(int)EnumYpprItem.POST_NUM] = "1";
                // フリー項目
                strYpprData[(int)EnumYpprItem.FREE_ITEM] = "3" + m_listOrderData[i].strOrderId;
                // 代引金額
                if (m_listOrderData[i].strPaymentMethod == "代金引換")
                {
                    strYpprData[(int)EnumYpprItem.COD_PAY] = m_listOrderData[i].strTotalPay;
                }
                else
                {
                    strYpprData[(int)EnumYpprItem.COD_PAY] = "";
                }
                // 代引消費税
                strYpprData[(int)EnumYpprItem.COD_TAX] = "";
                // 商品名設定
                strYpprData[(int)EnumYpprItem.PRODUCT_NAME] = "電子部品";
                // 商品備考
                strYpprData[(int)EnumYpprItem.PRODUCT_RENARKS] = "";


                // CSVへ書き込む
                StreamWriter clsSw;
                try
                {
                    clsSw = new StreamWriter(m_YpprFileName, true, Encoding.GetEncoding("Shift_JIS"));
                }
                catch
                {
                    return false;
                }
                // 代引きで複数箱の場合はコレクトと発払いに分ける
                if (m_listOrderData[i].strPaymentMethod == "代金引換")
                {
                    // コレクト
                    for (int j = 0; j < (int)EnumYpprItem.MAX; j++)
                    {
                        clsSw.Write("{0},", strYpprData[j]);
                    }
                    clsSw.Write("\n");

                    // 発払い
                    for (int j = 1; j < m_listOrderData[i].iTotalBoxNum; j++)
                    {
                        strYpprData[(int)EnumYpprItem.COD_PAY] = "";
                        strYpprData[(int)EnumYpprItem.PAYMENT_SOURCE] = "0";
                        strYpprData[(int)EnumYpprItem.INVOICE_CLASS] = "1100783001";
                        for (int k = 0; k < (int)EnumYpprItem.MAX; k++)
                        {
                            clsSw.Write("{0},", strYpprData[k]);
                        }
                        clsSw.Write("\n");
                    }
                }
                else
                {
                    for (int j = 0; j < m_listOrderData[i].iTotalBoxNum; j++)
                    {
                        for (int k = 0; k < (int)EnumYpprItem.MAX; k++)
                        {
                            clsSw.Write("{0},", strYpprData[k]);
                        }
                        clsSw.Write("\n");
                    }
                }

                clsSw.Flush();
                clsSw.Close();
            }

            return true;
        }
        //--------------------------------------------------------------
        void MakeProductData()
        {
            m_listProduct.Clear();

            for (int i = 0; i < m_listOrderData.Count; i++)
            {
                if (m_listOrderData[i].bEnable == false)
                {
                    continue;
                }
                for (int j = 0; j < m_listOrderData[i].listProductName.Count; j++)
                {
                    bool bMacheFlg = false;
                    for (int k = 0; k < m_listProduct.Count; k++)
                    {
                        if (m_listProduct[k].strProductName == m_listOrderData[i].listProductName[j])
                        {
                            m_listProduct[k].iTotalNum = m_listProduct[k].iTotalNum + m_listOrderData[i].listQuantity[j];
                            bMacheFlg = true;
                            break;
                        }
                    }
                    if (bMacheFlg == false)
                    {
                        CProductData clsData = new CProductData();

                        clsData.strProductName = m_listOrderData[i].listProductName[j];
                        clsData.iTotalNum = 1;

                        m_listProduct.Add(clsData);
                    }
                }
            }
            m_listProduct.Sort(CompareByID);
        }
        private static int CompareByID(CProductData a, CProductData b)
        {
            return String.Compare(a.strProductName, b.strProductName);
        }
        //--------------------------------------------------------------
        /// <summary>
        /// 書き込む
        /// </summary>
        public bool WriteProductData()
        {

            StreamWriter clsSw;
            try
            {
                clsSw = new StreamWriter(m_ListProductFileName, false, Encoding.GetEncoding("Shift_JIS"));
            }
            catch
            {
                return false;
            }

            // データ書き込み
            clsSw.Write("製品名,注文数\n");
            for (int i = 0; i < m_listProduct.Count; i++)
            {
                clsSw.Write("{0},{1}", m_listProduct[i].strProductName, m_listProduct[i].iTotalNum.ToString());
                clsSw.Write("\n");
            }


            clsSw.Flush();
            clsSw.Close();

            return true;
        }
        //--------------------------------------------------------------
    }
}
