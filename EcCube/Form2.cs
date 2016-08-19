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
        // ヤマト出力ファイル名
        string m_YamatoFileName;
        // Openlogi出力ファイル名
        string m_OpenlogiFileName;

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

            // ヤマト出力ファイル名
            m_YamatoFileName = Path.GetDirectoryName(strFileName) + "\\" + "ヤマト.csv";
            m_OpenlogiFileName = Path.GetDirectoryName(strFileName) + "\\" + "OpenLogi.csv";

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
                        clsEcCubeData.iTotalBoxNum += (3 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn006")
                    {
                        clsEcCubeData.iTotalBoxNum += (2 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn001")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn002")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn005")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn007")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn014")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn015")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn019")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn016")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn018")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn017")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "cfd002")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "cfd003")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "cfd004")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "cfd008")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "cfd011")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "cfd012")
                    {
                        clsEcCubeData.iTotalBoxNum += (1 * clsEcCubeData.listQuantity[i]);
                    }
                    else if (clsEcCubeData.listProductCode[i] == "mcn003")
                    {
                        // メガネはMiniの中に入れる
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
                else if (clsEcCubeData.iTotalBoxNum >= 3)
                {
                    clsEcCubeData.enInvoiceClass = EnumInvoiceClass.YAMATO_POST;
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
            // ヤマト便
            if (ConvertYamatoPost() == false)
            {
                return true;
            }
            // オープンロジ
            if (ConvertOpenlogi() == false)
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
                if (m_listOrderData[i].bEnable == false || m_listOrderData[i].enInvoiceClass == EnumInvoiceClass.YAMATO_POST)
                {
                    continue;
                }

                // ヤマトデーター
                string[] strYamatoData = new string[(int)EnumYamatoItem.MAX];

                // お客様側管理番号
                strYamatoData[(int)EnumYamatoItem.USER_ID] = "3" + m_listOrderData[i].strOrderId;
                // 送り状種別
                if (m_listOrderData[i].strPaymentMethod == "代金引換")
                {
                    strYamatoData[(int)EnumYamatoItem.INVOICE_CLASS] = "2";
                }
                else
                {
                    strYamatoData[(int)EnumYamatoItem.INVOICE_CLASS] = "0";
                }
                // 発送予定日
                strYamatoData[(int)EnumYamatoItem.SHIPPING_SCHEDULE_TIME] = DateTime.Now.ToString("yyyy/MM/dd");
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
                        strYamatoData[(int)EnumYamatoItem.DELIBERY_DATE] = m_listOrderData[i].strDeliberyDate.Substring(0, 4) +
                                                                       "/" +
                                                                       m_listOrderData[i].strDeliberyDate.Substring(5, 2) +
                                                                       "/" +
                                                                       m_listOrderData[i].strDeliberyDate.Substring(8, 2);
                    }
                    else
                    {
                        strYamatoData[(int)EnumYamatoItem.DELIBERY_DATE] = "";
                    }

                }
                else
                {
                    strYamatoData[(int)EnumYamatoItem.DELIBERY_DATE] = "";
                }
                // お届け時間
                if (m_listOrderData[i].strDeliberyTime == "午前中")
                {
                    strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "0812";
                }
                else if (m_listOrderData[i].strDeliberyTime == "12～14")
                {
                    strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "1214";
                }
                else if (m_listOrderData[i].strDeliberyTime == "14～16")
                {
                    strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "1416";
                }
                else if (m_listOrderData[i].strDeliberyTime == "16～18")
                {
                    strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "1618";
                }
                else if (m_listOrderData[i].strDeliberyTime == "18～20")
                {
                    strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "1820";
                }
                else
                {
                    strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "";
                }
                // お届け先電話番号
                strYamatoData[(int)EnumYamatoItem.TRANSPORT_TEL] = m_listOrderData[i].strPhoneNo;
                // お届け先郵便番号
                strYamatoData[(int)EnumYamatoItem.TRANSPORT_POST_NO] = m_listOrderData[i].strPostNo;
                // お届け先住所
                int iPos;

                iPos = m_listOrderData[i].strAddress3.IndexOf(" ");
                if (iPos <= 0)
                {
                    iPos = m_listOrderData[i].strAddress3.IndexOf("　");
                    if (iPos <= 0)
                    {
                        strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_1] = m_listOrderData[i].strAddress1 + m_listOrderData[i].strAddress2 + m_listOrderData[i].strAddress3;
                        strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_2] = "";
                    }
                    else
                    {
                        strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_1] = m_listOrderData[i].strAddress1 + m_listOrderData[i].strAddress2 + m_listOrderData[i].strAddress3.Substring(0, iPos);
                        strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_2] = m_listOrderData[i].strAddress3.Substring(iPos + 1, m_listOrderData[i].strAddress3.Length - (iPos + 1));
                    }
                }
                else
                {
                    strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_1] = m_listOrderData[i].strAddress1 + m_listOrderData[i].strAddress2 + m_listOrderData[i].strAddress3.Substring(0, iPos);
                    strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_2] = m_listOrderData[i].strAddress3.Substring(iPos + 1, m_listOrderData[i].strAddress3.Length - (iPos + 1));
                }

                // お届け先会社・部門名1
                strYamatoData[(int)EnumYamatoItem.TRANSPORT_COMPANY_1] = m_listOrderData[i].strCompanyName;
                // お届け先会社・部門名2
                strYamatoData[(int)EnumYamatoItem.TRANSPORT_COMPANY_2] = "";
                // お届け先名
                strYamatoData[(int)EnumYamatoItem.TRANSPORT_NAME] = m_listOrderData[i].strName;
                // お届け先敬称
                strYamatoData[(int)EnumYamatoItem.TRANSPORT_TITLE] = "様";
                // ご依頼主コード
                strYamatoData[(int)EnumYamatoItem.ORIGIN_CODE] = "";
                // ご依頼主電話番号
                strYamatoData[(int)EnumYamatoItem.ORIGIN_TEL] = "050-3786-7989";
                // ご依頼主郵便番号
                strYamatoData[(int)EnumYamatoItem.ORIGIN_POST_NO] = "4000306";
                // ご依頼主住所1
                strYamatoData[(int)EnumYamatoItem.ORIGIN_ADDRESS_1] = "山梨県南アルプス市小笠原1589-1";
                // ご依頼主住所2
                strYamatoData[(int)EnumYamatoItem.ORIGIN_ADDRESS_2] = "";
                // ご依頼主名
                strYamatoData[(int)EnumYamatoItem.ORIGIN_NAME] = "smartDIYs Shop";
                // 品名コード1
                strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_CODE_1] = "";
                // 品名1
                strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_1] = m_listOrderData[i].listProductName[0];
                if (strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_1].Length > 25)
                {
                    strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_1] = strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_1].Substring(0, 25);
                }
                // 品名コード2
                strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_CODE_2] = "";
                // 品名2
                strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_2] = "";
                // 荷扱い1
                strYamatoData[(int)EnumYamatoItem.FREIGHT_HANDLING_1] = "精密機器";
                // 荷扱い2
                strYamatoData[(int)EnumYamatoItem.FREIGHT_HANDLING_2] = "下積厳禁";
                // 記事
                strYamatoData[(int)EnumYamatoItem.ARTICLE] = "";
                // 代引金額
                if (m_listOrderData[i].strPaymentMethod == "代金引換")
                {
                    strYamatoData[(int)EnumYamatoItem.COD_PAY] = m_listOrderData[i].strTotalPay;
                }
                else
                {
                    strYamatoData[(int)EnumYamatoItem.COD_PAY] = "";
                }
                // 代引消費税
                strYamatoData[(int)EnumYamatoItem.COD_TAX] = "";
                // 発行枚数
                strYamatoData[(int)EnumYamatoItem.POST_NUM] = m_listOrderData[i].iTotalBoxNum.ToString();
                // 個数口枠の印字
                strYamatoData[(int)EnumYamatoItem.NUMBER_FRAME] = "3";
                // ご請求先顧客コード
                strYamatoData[(int)EnumYamatoItem.BILLING_CODE] = "05037867989";
                // 運賃管理番号
                strYamatoData[(int)EnumYamatoItem.FARE_NO] = "01";
                // 空白
                strYamatoData[(int)EnumYamatoItem.NULL_1] = "";
                strYamatoData[(int)EnumYamatoItem.NULL_2] = "";
                strYamatoData[(int)EnumYamatoItem.NULL_3] = "";
                strYamatoData[(int)EnumYamatoItem.NULL_4] = "";
                strYamatoData[(int)EnumYamatoItem.NULL_5] = "";
                strYamatoData[(int)EnumYamatoItem.NULL_6] = "";
                strYamatoData[(int)EnumYamatoItem.NULL_7] = "";
                strYamatoData[(int)EnumYamatoItem.NULL_8] = "";
                strYamatoData[(int)EnumYamatoItem.NULL_9] = "";
                strYamatoData[(int)EnumYamatoItem.NULL_10] = "";

                // CSVへ書き込む
                bool bHeaderFlg = false;

                // ファイルが存在しない場合には、ヘッダーを追加
                if (System.IO.File.Exists(m_YamatoFileName) == false)
                {
                    bHeaderFlg = true;
                }

                StreamWriter clsSw;
                try
                {
                    clsSw = new StreamWriter(m_YamatoFileName, true, Encoding.GetEncoding("Shift_JIS"));
                }
                catch
                {
                    return false;
                }
                // ヘッダー書き込み
                if (bHeaderFlg)
                {
                    clsSw.Write("ヤマト運輸データ\n");
                }
                // 代引きで複数箱の場合はコレクトと発払いに分ける
                if (m_listOrderData[i].strPaymentMethod == "代金引換" && m_listOrderData[i].iTotalBoxNum > 1)
                {
                    // コレクト
                    strYamatoData[(int)EnumYamatoItem.POST_NUM] = "1";
                    for (int j = 0; j < (int)EnumYamatoItem.MAX; j++)
                    {
                        clsSw.Write("{0},", strYamatoData[j]);
                    }
                    clsSw.Write("\n");
                
                    // 発払い                
                    strYamatoData[(int)EnumYamatoItem.POST_NUM] = (m_listOrderData[i].iTotalBoxNum - 1).ToString();
                    strYamatoData[(int)EnumYamatoItem.COD_PAY] = "";
                    strYamatoData[(int)EnumYamatoItem.INVOICE_CLASS] = "0";
                    for (int j = 0; j < (int)EnumYamatoItem.MAX; j++)
                    {
                        clsSw.Write("{0},", strYamatoData[j]);
                    }
                    clsSw.Write("\n");
                }
                else
                {
                    for (int j = 0; j < (int)EnumYamatoItem.MAX; j++)
                    {
                        clsSw.Write("{0},", strYamatoData[j]);
                    }
                    clsSw.Write("\n");
                }

                clsSw.Flush();
                clsSw.Close();
            }

            return true;
        }
        //--------------------------------------------------------------
        public bool ConvertYamatoPost()
        {
            //ドキュメントを作成
            const string strPdfFileName = "LabelPrintTemp.pdf";
            Document doc = new Document(PageSize.A4, 10f, 10f, 25f, 0);
            bool bPDFShowFlg = false;

            try
            {
                //ファイルの出力先を設定
                PdfWriter.GetInstance(doc, new FileStream(strPdfFileName, FileMode.Create));
                //ドキュメントを開く
                doc.Open();


                for (int i = 0; i < m_listOrderData.Count; i++)
                {
                    if (m_listOrderData[i].bEnable == false || m_listOrderData[i].enInvoiceClass != EnumInvoiceClass.YAMATO_POST)
                    {
                        continue;
                    }
                    if (bPDFShowFlg)
                    {
                        // 改ページ
                        doc.NewPage();
                    }

                    bPDFShowFlg = true;

                    float[] headerwodth = new float[] { 1f, 1f };
                    PdfPTable tbl = new PdfPTable(headerwodth);
                    tbl.WidthPercentage = 100;

                    int iOutCnt = 0;
                    
                    for (int j = 0; j < m_listOrderData[i].iTotalBoxNum; j++)
                    {
                        if (MakeLabelPrint(tbl, m_listOrderData[i], j + 1) == false)
                        {
                            return false;
                        }
                        iOutCnt++;
                    }

                    // 空白を埋める
                    for (int j = 0; j < headerwodth.Length - (iOutCnt % headerwodth.Length); j++)
                    {
                        SpaceLabelPrint(tbl);
                    }

                    // テーブル追加
                    doc.Add(tbl);
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                if (bPDFShowFlg)
                {
                    //ドキュメントを閉じる
                    doc.Close();
                }
            }

            if (bPDFShowFlg)
            {
                // PDFを開く
                System.Diagnostics.Process.Start(strPdfFileName);
            }

            return true;
        }
        //--------------------------------------------------------------
        public bool ConvertOpenlogi()
        {
            for (int i = 0; i < m_listOrderData.Count; i++)
            {
                if (m_listOrderData[i].bEnable == false)
                {
                    continue;
                }

                for (int j = 0; j < m_listOrderData[i].listProductCode.Count; j++)
                {
                    if (m_listOrderData[i].listProductCode[j] != "mcn004")
                    {
                        continue;
                    }

                    // OpenLogiデーター
                    string[] strOpenlogiData = new string[(int)EnumOpenLogiItem.MAX];

                    // 注文番号
                    strOpenlogiData[(int)EnumOpenLogiItem.ORDER_ID] = "3" + m_listOrderData[i].strOrderId;
                    // 出庫依頼数
                    strOpenlogiData[(int)EnumOpenLogiItem.ORDER_NUM] = m_listOrderData[i].listQuantity[j].ToString();
                    // お届け先郵便番号
                    strOpenlogiData[(int)EnumOpenLogiItem.TRANSPORT_POST_NO] = m_listOrderData[i].strPostNo;
                    // お届け先住所1
                    strOpenlogiData[(int)EnumOpenLogiItem.TRANSPORT_ADDRESS_1] = m_listOrderData[i].strAddress1;
                    // お届け先住所2
                    strOpenlogiData[(int)EnumOpenLogiItem.TRANSPORT_ADDRESS_2] = m_listOrderData[i].strAddress2 + m_listOrderData[i].strAddress3;
                    // お届け先住所3
                    strOpenlogiData[(int)EnumOpenLogiItem.TRANSPORT_ADDRESS_3] = "";
                    // お届け先名
                    strOpenlogiData[(int)EnumOpenLogiItem.TRANSPORT_NAME] = m_listOrderData[i].strName;
                    // お届け先会社・部門名
                    strOpenlogiData[(int)EnumOpenLogiItem.TRANSPORT_COMPANY] = m_listOrderData[i].strCompanyName;
                    // お届け先電話番号
                    strOpenlogiData[(int)EnumOpenLogiItem.TRANSPORT_TEL] = m_listOrderData[i].strPhoneNo;
                    // 割れ物注意
                    strOpenlogiData[(int)EnumOpenLogiItem.FRAGILE] = "1";

                    // ご依頼主郵便番号
                    strOpenlogiData[(int)EnumOpenLogiItem.ORIGIN_POST_NO] = "4000306";
                    // ご依頼主住所1
                    strOpenlogiData[(int)EnumOpenLogiItem.ORIGIN_ADDRESS_1] = "山梨県";
                    // ご依頼主住所2
                    strOpenlogiData[(int)EnumOpenLogiItem.ORIGIN_ADDRESS_2] = "南アルプス市小笠原1589-1";
                    // ご依頼主住所3
                    strOpenlogiData[(int)EnumOpenLogiItem.ORIGIN_ADDRESS_3] = "";
                    // ご依頼主名
                    strOpenlogiData[(int)EnumOpenLogiItem.ORIGIN_NAME] = "有井";
                    // ご依頼主会社・部門名
                    strOpenlogiData[(int)EnumOpenLogiItem.ORIGIN_COMPANY] = "株式会社smartDIYs";
                    // ご依頼主電話番号
                    strOpenlogiData[(int)EnumOpenLogiItem.ORIGIN_TEL] = "050-3786-7989";


                    // CSVへ書き込む
                    bool bHeaderFlg = false;

                    // ファイルが存在しない場合には、ヘッダーを追加
                    if (System.IO.File.Exists(m_OpenlogiFileName) == false)
                    {
                        bHeaderFlg = true;
                    }

                    StreamWriter clsSw;
                    try
                    {
                        clsSw = new StreamWriter(m_OpenlogiFileName, true, Encoding.GetEncoding("Shift_JIS"));
                    }
                    catch
                    {
                        return false;
                    }
                    // ヘッダー書き込み
                    if (bHeaderFlg)
                    {
                        clsSw.Write("ID,商品名,商品コード,サイズ,在庫数,出庫依頼数,出庫予約,単価,合計金額,小計,配送料,手数料,割引額,総計,配送先郵便番号,配送先都道府県,配送先住所,配送先マンション・ビル名,配送先氏名,配送先会社名,配送先電話番号,配送先連絡メール,配送先マスタ登録,注文番号,配送便指定,配送会社指定,希望時間帯,お届け希望日,ギフトラッピング単位,ラッピングタイプ,贈り主氏名,同梱指定,不在時宅配ボックス,到着前電話確認,割れ物注意,代引き指定,明細注意書き,ご依頼主郵便番号,ご依頼主都道府県,ご依頼主住所,ご依頼主マンション・ビル名,ご依頼主氏名,ご依頼主会社名,ご依頼主部署名,ご依頼主電話番号\n");
                    }

                    // 板金
                    // 商品管理コード
                    strOpenlogiData[(int)EnumOpenLogiItem.PRODUCT_ID] = "43";
                    // 商品名
                    strOpenlogiData[(int)EnumOpenLogiItem.PRODUCT_NAME] = "Smart Laser CO2 Cover(co2-160311-100)";
                    // 商品コード
                    strOpenlogiData[(int)EnumOpenLogiItem.PRODUCT_CODE] = "7";
                    for (int k = 0; k < (int)EnumOpenLogiItem.MAX; k++)
                    {
                        clsSw.Write("{0}", strOpenlogiData[k]);
                        if (k != (int)EnumOpenLogiItem.MAX - 1)
                        {
                            clsSw.Write(",");
                        }
                    }
                    clsSw.Write("\n");

                    // アクリル
                    // 商品管理コード
                    strOpenlogiData[(int)EnumOpenLogiItem.PRODUCT_ID] = "44";
                    // 商品名
                    strOpenlogiData[(int)EnumOpenLogiItem.PRODUCT_NAME] = "Smart Laser CO2 Acrylic(co2-160311-100)";
                    // 商品コード
                    strOpenlogiData[(int)EnumOpenLogiItem.PRODUCT_CODE] = "8";
                    for (int k = 0; k < (int)EnumOpenLogiItem.MAX; k++)
                    {
                        clsSw.Write("{0}", strOpenlogiData[k]);
                        if (k != (int)EnumOpenLogiItem.MAX - 1)
                        {
                            clsSw.Write(",");
                        }
                    }
                    clsSw.Write("\n");

                    // アルミフレーム
                    // 商品管理コード
                    strOpenlogiData[(int)EnumOpenLogiItem.PRODUCT_ID] = "42";
                    // 商品名
                    strOpenlogiData[(int)EnumOpenLogiItem.PRODUCT_NAME] = "Smart Laser CO2 Aluminum Extrusion(co2-160311-100)";
                    // 商品コード
                    strOpenlogiData[(int)EnumOpenLogiItem.PRODUCT_CODE] = "6";
                    for (int k = 0; k < (int)EnumOpenLogiItem.MAX; k++)
                    {
                        clsSw.Write("{0}", strOpenlogiData[k]);
                        if (k != (int)EnumOpenLogiItem.MAX - 1)
                        {
                            clsSw.Write(",");
                        }
                    }
                    clsSw.Write("\n");

                    clsSw.Flush();
                    clsSw.Close();
                }
            }

            return true;
        }
        //--------------------------------------------------------------
        /// <summary>
        /// ラベル空白設定
        /// </summary>
        public bool SpaceLabelPrint(PdfPTable tbl)
        {
            Font fnt;
            PdfPCell cel;

            fnt = new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 22);
            cel = new PdfPCell(new Phrase("", fnt));
            cel.HorizontalAlignment = Element.ALIGN_LEFT;
            cel.FixedHeight = 199f;
            cel.Padding = 7;
            cel.Border = 0;
            tbl.AddCell(cel);

            return true;
        }
        //--------------------------------------------------------------
        /// <summary>
        /// ラベル作成
        /// </summary>
        public bool MakeLabelPrint(PdfPTable tbl, CEcCubeData clsEcCubeData, int iPos)
        {
            Font fnt;
            PdfPCell cel;
            string strBuf = iPos.ToString() + "/" + clsEcCubeData.iTotalBoxNum.ToString() + "\n〒" + clsEcCubeData.strPostNo + "\n" + clsEcCubeData.strAddress1 + clsEcCubeData.strAddress2 + clsEcCubeData.strAddress3 + "\n" + clsEcCubeData.strCompanyName + " " + clsEcCubeData.strName + "様" + "\n" + "TEL：" + clsEcCubeData.strPhoneNo;

            fnt = new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 22);
            cel = new PdfPCell(new Phrase(strBuf, fnt));
            cel.HorizontalAlignment = Element.ALIGN_LEFT;
            cel.FixedHeight = 199f;
            cel.Padding = 7;
            cel.Border = 0;
            tbl.AddCell(cel);

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
