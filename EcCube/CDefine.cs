using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EcCube
{
    //--------------------------------------------------------------
    #region 列挙型
    /// <summary>
    /// 送り状種別要素
    /// </summary>
    public enum EnumInvoiceClass
    {
        /// <summary>宅急便</summary>
        DELIVERY = 0,
        /// <summary>代引き</summary>
        CASH_ON_DELIVERY,
        /// <summary>ヤマト便</summary>
        YAMATO_POST,

        MAX
    }
    /// <summary>
    /// EC-CUBE要素
    /// </summary>
    public enum EnumEcCubeItem
    {
        /// <summary>注文番号</summary>
        ORDER_ID = 0,
        /// <summary>姓</summary>
        FIRST_NAME,
        /// <summary>名</summary>
        LAST_NAME,
        /// <summary>郵便番号1</summary>
        POST_NO_1,
        /// <summary>郵便番号2</summary>
        POST_NO_2,
        /// <summary>住所1</summary>
        ADDRESS_1,
        /// <summary>住所2</summary>
        ADDRESS_2,
        /// <summary>住所3</summary>
        ADDRESS_3,
        /// <summary>会社名</summary>
        COMPANY_NAME,
        /// <summary>電話番号1</summary>
        PHONE_NO_1,
        /// <summary>電話番号2</summary>
        PHONE_NO_2,
        /// <summary>電話番号3</summary>
        PHONE_NO_3,
        /// <summary>お届け日</summary>
        DELIBERY_DATE,
        /// <summary>お届け時間</summary>
        DELIBERY_TIME,
        /// <summary>支払総額</summary>
        TOTAL_PAY,
        /// <summary>お支払方法</summary>
        PAYMENT_METHOD,
        /// <summary>配送方法</summary>
        DELIVERY_METHOD,
        /// <summary>メールアドレス</summary>
        E_MAIL,
        /// <summary>商品コード</summary>
        PRODUCT_CODE,
        /// <summary>数量</summary>
        QUANTITY,
        /// <summary>商品名</summary>
        PRODUCT_NAME,

        MAX
    }
    /// <summary>
    /// ヤマト要素
    /// </summary>
    public enum EnumYamatoItem
    {
        /// <summary>お客様側管理番号</summary>
        USER_ID = 0,
        /// <summary>送り状種別</summary>
        INVOICE_CLASS,
        /// <summary>NULL</summary>
        NULL_1,
        /// <summary>NULL</summary>
        NULL_2,
        /// <summary>発送予定時間</summary>
        SHIPPING_SCHEDULE_TIME,
        /// <summary>配達指定日</summary>
        DELIBERY_DATE,
        /// <summary>配達時間帯区分</summary>
        DELIBERY_TIME,
        /// <summary>NULL</summary>
        NULL_3,
        /// <summary>お届け先電話番号</summary>
        TRANSPORT_TEL,
        /// <summary>NULL</summary>
        NULL_4,
        /// <summary>お届け先郵便番号</summary>
        TRANSPORT_POST_NO,
        /// <summary>お届け先住所1</summary>
        TRANSPORT_ADDRESS_1,
        /// <summary>お届け先住所2</summary>
        TRANSPORT_ADDRESS_2,
        /// <summary>お届け先会社・部門名1</summary>
        TRANSPORT_COMPANY_1,
        /// <summary>お届け先会社・部門名2</summary>
        TRANSPORT_COMPANY_2,
        /// <summary>お届け先名</summary>
        TRANSPORT_NAME,
        /// <summary>NULL</summary>
        NULL_5,
        /// <summary>お届け先敬称</summary>
        TRANSPORT_TITLE,
        /// <summary>ご依頼主コード</summary>
        ORIGIN_CODE,
        /// <summary>ご依頼主電話番号</summary>
        ORIGIN_TEL,
        /// <summary>NULL</summary>
        NULL_6,
        /// <summary>ご依頼主郵便番号</summary>
        ORIGIN_POST_NO,
        /// <summary>ご依頼主住所1</summary>
        ORIGIN_ADDRESS_1,
        /// <summary>ご依頼主住所2</summary>
        ORIGIN_ADDRESS_2,
        /// <summary>ご依頼主名</summary>
        ORIGIN_NAME,
        /// <summary>NULL</summary>
        NULL_7,
        /// <summary>品名コード1</summary>
        PRODUCT_NAME_CODE_1,
        /// <summary>品名1</summary>
        PRODUCT_NAME_1,
        /// <summary>品名コード2</summary>
        PRODUCT_NAME_CODE_2,
        /// <summary>品名2</summary>
        PRODUCT_NAME_2,
        /// <summary>荷扱い1</summary>
        FREIGHT_HANDLING_1,
        /// <summary>荷扱い2</summary>
        FREIGHT_HANDLING_2,
        /// <summary>記事</summary>
        ARTICLE,
        /// <summary>代引金額</summary>
        COD_PAY,
        /// <summary>代引消費税</summary>
        COD_TAX,
        /// <summary>NULL</summary>
        NULL_8,
        /// <summary>NULL</summary>
        NULL_9,
        /// <summary>発行枚数</summary>
        POST_NUM,
        /// <summary>個数口枠の印字</summary>
        NUMBER_FRAME,
        /// <summary>ご請求先顧客コード</summary>
        BILLING_CODE,
        /// <summary>NULL</summary>
        NULL_10,
        /// <summary>運賃管理番号</summary>
        FARE_NO,

        MAX
    }
    /// <summary>
    /// OpenLogi要素
    /// </summary>
    public enum EnumOpenLogiItem
    {
        /// <summary>商品管理コード</summary>
        PRODUCT_ID = 0,
        /// <summary>商品名</summary>
        PRODUCT_NAME,
        /// <summary>商品コード</summary>
        PRODUCT_CODE,
        /// <summary>NULL</summary>
        NULL_1,
        /// <summary>NULL</summary>
        NULL_2,
        /// <summary>出庫依頼数</summary>
        ORDER_NUM,
        /// <summary>NULL</summary>
        NULL_3,
        /// <summary>NULL</summary>
        NULL_4,
        /// <summary>NULL</summary>
        NULL_5,
        /// <summary>NULL</summary>
        NULL_6,
        /// <summary>NULL</summary>
        NULL_7,
        /// <summary>NULL</summary>
        NULL_8,
        /// <9ummary>NULL</summary>
        NULL_10,
        /// <summary>NULL</summary>
        NULL_11,
        /// <summary>お届け先郵便番号</summary>
        TRANSPORT_POST_NO,
        /// <summary>お届け先住所1</summary>
        TRANSPORT_ADDRESS_1,
        /// <summary>お届け先住所2</summary>
        TRANSPORT_ADDRESS_2,
        /// <summary>お届け先住所3</summary>
        TRANSPORT_ADDRESS_3,
        /// <summary>お届け先名</summary>
        TRANSPORT_NAME,
        /// <summary>お届け先会社・部門名</summary>
        TRANSPORT_COMPANY,
        /// <summary>お届け先電話番号</summary>
        TRANSPORT_TEL,
        /// <summary>NULL</summary>
        NULL_12,
        /// <summary>NULL</summary>
        NULL_13,
        /// <summary>注文番号</summary>
        ORDER_ID,
        /// <summary>NULL</summary>
        NULL_14,
        /// <summary>NULL</summary>
        NULL_15,
        /// <summary>NULL</summary>
        NULL_16,
        /// <summary>NULL</summary>
        NULL_17,
        /// <summary>NULL</summary>
        NULL_18,
        /// <summary>NULL</summary>
        NULL_19,
        /// <summary>NULL</summary>
        NULL_20,
        /// <summary>NULL</summary>
        NULL_21,
        /// <summary>NULL</summary>
        NULL_22,
        /// <summary>NULL</summary>
        NULL_23,
        /// <summary>割れ物注意</summary>
        FRAGILE,
        /// <summary>NULL</summary>
        NULL_24,
        /// <summary>NULL</summary>
        NULL_25,
        /// <summary>ご依頼主郵便番号</summary>
        ORIGIN_POST_NO,
        /// <summary>ご依頼主住所1</summary>
        ORIGIN_ADDRESS_1,
        /// <summary>ご依頼主住所2</summary>
        ORIGIN_ADDRESS_2,
        /// <summary>ご依頼主住所3</summary>
        ORIGIN_ADDRESS_3,
        /// <summary>ご依頼主名</summary>
        ORIGIN_NAME,
        /// <summary>ご依頼主会社・部門名</summary>
        ORIGIN_COMPANY,
        /// <summary>NULL</summary>
        NULL_26,
        /// <summary>ご依頼主電話番号</summary>
        ORIGIN_TEL,

        MAX
    }
    #endregion
    //--------------------------------------------------------------
    #region 構造体
    /// <summary>
    /// EC-CUBEデータクラス
    /// </summary>
    public class CEcCubeData
    {
        //--------------------------------------------------------------
        /// <summary>有効無効Flg</summary>
        public bool bEnable;
        /// <summary>注文番号</summary>
        public string strOrderId;
        /// <summary>購入者名</summary>
        public string strName;
        /// <summary>郵便番号</summary>
        public string strPostNo;
        /// <summary>住所1</summary>
        public string strAddress1;
        /// <summary>住所2</summary>
        public string strAddress2;
        /// <summary>住所3</summary>
        public string strAddress3;
        /// <summary>会社名</summary>
        public string strCompanyName;
        /// <summary>電話番号</summary>
        public string strPhoneNo;
        /// <summary>お届け日</summary>
        public string strDeliberyDate;
        /// <summary>お届け時間</summary>
        public string strDeliberyTime;
        /// <summary>支払総額</summary>
        public string strTotalPay;
        /// <summary>お支払方法</summary>
        public string strPaymentMethod;
        /// <summary>配送方法</summary>
        public string strDeliveryMethod;
        /// <summary>メールアドレス</summary>
        public string strEMail;
        /// <summary>商品名コード</summary>
        public List<string> listProductCode = new List<string>();
        /// <summary>数量</summary>
        public List<int> listQuantity = new List<int>();
        /// <summary>商品名リスト</summary>
        public List<string> listProductName = new List<string>();
        /// <summary>送り状種別</summary>
        public EnumInvoiceClass enInvoiceClass;
        /// <summary>合計箱数</summary>
        public int iTotalBoxNum;



        //--------------------------------------------------------------
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public CEcCubeData()
        {
            listProductCode = new List<string>();
            listQuantity = new List<int>();
            listProductName = new List<string>();
            Init();
        }
        /// <summary>
        /// 初期化関数
        /// </summary>
        public void Init()
        {
            bEnable = false;
            strOrderId = "";
            strName = "";
            strPostNo = "";
            strAddress1 = "";
            strAddress2 = "";
            strAddress3 = "";
            strCompanyName = "";
            strPhoneNo = "";
            strDeliberyDate = "";
            strDeliberyTime = "";
            strTotalPay = "";
            strPaymentMethod = "";
            strDeliveryMethod = "";
            strEMail = "";
            listProductCode.Clear();
            listQuantity.Clear();
            listProductName.Clear();
            enInvoiceClass = EnumInvoiceClass.DELIVERY;
            iTotalBoxNum = 0;
        }
    }
    /// <summary>
    /// 製品データクラス
    /// </summary>
    public class CProductData
    {
        //--------------------------------------------------------------
        /// <summary>製品名</summary>
        public string strProductName;
        /// <summary>合計数</summary>
        public int iTotalNum;



        //--------------------------------------------------------------
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public CProductData()
        {
            Init();
        }
        /// <summary>
        /// 初期化関数
        /// </summary>
        public void Init()
        {
            strProductName = "";
            iTotalNum = 0;
        }
    }
    #endregion
    //--------------------------------------------------------------
    /// <summary>
    /// 定数宣言クラス
    /// </summary>
    class CDefine
    {
    }
    //--------------------------------------------------------------
}
