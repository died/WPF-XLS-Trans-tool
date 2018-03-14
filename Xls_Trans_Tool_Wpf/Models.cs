using System.Collections.Generic;

namespace Xls_Trans_Tool_Wpf
{
    public class Models
    {
        /// <summary>
        /// full adidas 20171228 new version
        /// </summary>
        public static List<Column> AdidasColumnList = new List<Column>
        {
            new Column {Name = "Issue Date:", Field = "A", DateTimeColumn = true},
            new Column {Name = "Order Number:", Field = "B", DateTimeColumn = false},
            new Column {Name = "Latest Update (GMT):", Field = "C", DateTimeColumn = true},
            new Column {Name = "Version:", Field = "D", DateTimeColumn = false},
            new Column {Name = "Purpose:", Field = "E", DateTimeColumn = false},
            new Column {Name = "Buyer (T1 Supplier Name)", Field = "F", DateTimeColumn = false},
            new Column {Name = "T1 address line 1", Field = "G", DateTimeColumn = false},
            new Column {Name = "T1 address line 2", Field = "H", DateTimeColumn = false},
            new Column {Name = "T1 address line 3", Field = "I", DateTimeColumn = false},
            new Column {Name = "T1 address line 4", Field = "J", DateTimeColumn = false},
            new Column {Name = "T1 address city", Field = "K", DateTimeColumn = false},
            new Column {Name = "T1 address province/ state", Field = "L", DateTimeColumn = false},
            new Column {Name = "T1 address postal code", Field = "M", DateTimeColumn = false},
            new Column {Name = "T1 address Country (refer to Pick List_Country)" ,Field = "N" ,DateTimeColumn = false},
            new Column {Name = "Contact Person (Buyer)" ,Field = "O" ,DateTimeColumn = false},
            new Column {Name = "Contact No. (Buyer)" ,Field = "P" ,DateTimeColumn = false},
            new Column {Name = "Email (Buyer)" ,Field = "Q" ,DateTimeColumn = false},
            new Column {Name = "Ship To (T1 Factory Name)" ,Field = "R" ,DateTimeColumn = false},
            new Column {Name = "T1 address line 1" ,Field = "S" ,DateTimeColumn = false},
            new Column {Name = "T1 address line 2" ,Field = "T" ,DateTimeColumn = false},
            new Column {Name = "T1 address line 3" ,Field = "U" ,DateTimeColumn = false},
            new Column {Name = "T1 address line 4" ,Field = "V" ,DateTimeColumn = false},
            new Column {Name = "T1 address city" ,Field = "W" ,DateTimeColumn = false},
            new Column {Name = "T1 address province/ state" ,Field = "X" ,DateTimeColumn = false},
            new Column {Name = "T1 address postal code", Field = "Y" ,DateTimeColumn = false},
            new Column {Name = "T1 address Country (refer to Pick List_Country)", Field = "Z" ,DateTimeColumn = false},
            new Column {Name = "T1 6 digit adidas Factory Code", Field = "AA", DateTimeColumn = false},
            new Column {Name = "Seller (T2 Supplier Name)", Field = "AB", DateTimeColumn = false},
            new Column {Name = "Actual Manufacturer (T2 Factory Name)", Field = "AC", DateTimeColumn = false},
            new Column {Name = "T2 address line 1", Field = "AD", DateTimeColumn = false},
            new Column {Name = "T2 address line 2", Field = "AE", DateTimeColumn = false},
            new Column {Name = "T2 address line 3", Field = "AF", DateTimeColumn = false},
            new Column {Name = "T2 address line 4", Field = "AG", DateTimeColumn = false},
            new Column {Name = "T2 address city", Field = "AH", DateTimeColumn = false},
            new Column {Name = "T2 address province/ state", Field = "AI", DateTimeColumn = false},
            new Column {Name = "T2 address postal code", Field = "AJ", DateTimeColumn = false},
            new Column {Name = "T2 address Country (refer to Pick List_Country)", Field = "AK", DateTimeColumn = false},
            new Column {Name = "T2 6 digit adidas Factory Code", Field = "AL", DateTimeColumn = false},
            new Column {Name = "T1 Customer", Field = "AM", DateTimeColumn = false},
            new Column {Name = "PO Download Date" ,Field = "AN" ,DateTimeColumn = true},
            new Column {Name = "Currency" ,Field = "AO" ,DateTimeColumn = false},
            new Column {Name = "Payment Terms" ,Field = "AP" ,DateTimeColumn = false},
            new Column {Name = "Incoterm" ,Field = "AQ" ,DateTimeColumn = false},
            new Column {Name = "Ship Mode" ,Field = "AR" ,DateTimeColumn = false},
            new Column {Name = "Country of Origin" ,Field = "AS" ,DateTimeColumn = false},
            new Column {Name = "Forwarder" ,Field = "AT" ,DateTimeColumn = false},
            new Column {Name = "Remarks" ,Field = "AU" ,DateTimeColumn = false},
            new Column {Name = "Packing Instruction:" ,Field = "AV" ,DateTimeColumn = false},
            new Column {Name = "Shipping Instruction:" ,Field = "AW" ,DateTimeColumn = false},
            new Column {Name = "Line Status" ,Field = "AX" ,DateTimeColumn = false},
            new Column {Name = "Line #", Field = "AY" ,DateTimeColumn = false},
            new Column {Name = "Ref#", Field = "AZ" ,DateTimeColumn = false},
            new Column {Name = "Description / Supplier Material Name", Field = "BA", DateTimeColumn = false},
            new Column {Name = "Width", Field = "BB", DateTimeColumn = false},
            new Column {Name = "UOM (for width)", Field = "BC", DateTimeColumn = false},
            new Column {Name = "Weight", Field = "BD", DateTimeColumn = false},
            new Column {Name = "UOM (for weight)", Field = "BE", DateTimeColumn = false},
            new Column {Name = "Length", Field = "BF", DateTimeColumn = false},
            new Column {Name = "UOM (for length)", Field = "BG", DateTimeColumn = false},
            new Column {Name = "Height", Field = "BH", DateTimeColumn = false},
            new Column {Name = "UOM (for Height)", Field = "BI", DateTimeColumn = false},
            new Column {Name = "Thickness", Field = "BJ", DateTimeColumn = false},
            new Column {Name = "UOM (for thickness)", Field = "BK", DateTimeColumn = false},
            new Column {Name = "Size", Field = "BL", DateTimeColumn = false},
            new Column {Name = "Material Color", Field = "BM", DateTimeColumn = false},
            new Column {Name = "Unit Price" ,Field = "BN" ,DateTimeColumn = false},
            new Column {Name = "Order Quantity" ,Field = "BO" ,DateTimeColumn = false},
            new Column {Name = "UOM (for order quantity)" ,Field = "BP" ,DateTimeColumn = false},
            new Column {Name = "Color Matching" ,Field = "BQ" ,DateTimeColumn = false},
            new Column {Name = "Section Name" ,Field = "BR" ,DateTimeColumn = false},
            new Column {Name = "Sustainable material" ,Field = "BS" ,DateTimeColumn = false},
            new Column {Name = "Buyer Request Date" ,Field = "BT" ,DateTimeColumn = true},
            new Column {Name = "adidas CRD" ,Field = "BU" ,DateTimeColumn = true},
            new Column {Name = "adidas Plan Date" ,Field = "BV" ,DateTimeColumn = true},
            new Column {Name = "Seller Confirm Delivery Date" ,Field = "BW" ,DateTimeColumn = true},
            new Column {Name = "Seller Updated Delivery Date" ,Field = "BX" ,DateTimeColumn = true},
            new Column {Name = "Confirmed Delivery Quantity", Field = "BY" ,DateTimeColumn = false},
            new Column {Name = "Updated Delivery Quantity", Field = "BZ" ,DateTimeColumn = false},
            new Column {Name = "Seller Confirm Delivery Date (Last Shipment)", Field = "CA", DateTimeColumn = true},
            new Column {Name = "Seller Updated Delivery Date(Last Shipment)", Field = "CB", DateTimeColumn = true},
            new Column {Name = "Confirmed Delivery Quantity(Last Shipment)", Field = "CC", DateTimeColumn = false},
            new Column {Name = "Updated Delivery Quantity(Last Shipment)", Field = "CD", DateTimeColumn = false},
            new Column {Name = "Season", Field = "CE", DateTimeColumn = false},
            new Column {Name = "Priority & Order Type", Field = "CF", DateTimeColumn = false},
            new Column {Name = "adidas Order Number", Field = "CG", DateTimeColumn = false},
            new Column {Name = "adidas Article Number", Field = "CH", DateTimeColumn = false},
            new Column {Name = "adidas Working Number / Model Name", Field = "CI", DateTimeColumn = false},
            new Column {Name = "Remarks", Field = "CJ", DateTimeColumn = false},
            new Column {Name = "Additional Optional 1", Field = "CK", DateTimeColumn = false},
            new Column {Name = "Additional Optional 2", Field = "CL", DateTimeColumn = false},
            new Column {Name = "Additional Optional 3", Field = "CM", DateTimeColumn = false},
            new Column {Name = "Additional Optional 4" ,Field = "CN" ,DateTimeColumn = false},
            new Column {Name = "Additional Optional 5" ,Field = "CO" ,DateTimeColumn = false}
        };

        //public static string[] SummaryHeaders =
        //{   //A                                         //E                                      //J                                                 
        //    "FTY", "STYLENO","Order Number","STOCKCODE","STOCKCOLOR","數量","單位", "價錢","季度","PI DATE","GMT DATE","Color Matching", "備註", "備註", "備註", "備註", "備註"
        //};
        //public static string[] SummaryHeaders =
        //{   //A                                                                                          //E
        //    "Ship To\n(T1\nFactory\nName)", "adidas Working\nNumber /\nModel Name","Order Number","Ref#","Material\nColor",
        //    //F                                                                //J                               //L
        //    "Order\nQuantity","UOM\n(for\norder\nquantity)", "Unit\nPrice","Season","Buyer\nRequest\nDate","adidas\nCRD","Color\nMatching","Shipping\nInstruction",
        //    //N
        //    "Remarks","Additional\nOptional 1", "Additional\nOptional 2", "Additional\nOptional 3", "Additional\nOptional 4", "Additional\nOptional 5"
        //};

        //public static List<string> SummaryHeaders = new List<string>
        //{
        //    //A                                                                                          //E
        //    "Ship To\n(T1\nFactory\nName)", "adidas Working\nNumber /\nModel Name","Order Number","Ref#","Material\nColor",
        //    //F                                                                     //J                                  //L
        //    "Order\nQuantity","UOM\n(for\norder\nquantity)", "Unit\nPrice","Season","Buyer\nRequest\nDate","adidas\nCRD","Seller Confirm\nDelivery Date"
        //    //M
        //    ,"Color\nMatching"//,"Shipping\nInstruction",
        //};

        public static List<string> SummaryHeaders = new List<string>
        {
            //A                                                                                          //E
            "Ship To\n(T1\nFactory\nName)", "adidas Working\nNumber /\nModel Name","Order Number","Ref#","Color\nCode","Color\nDesc",
            //G                                                                     //K                                 //M
            "Order\nQuantity","UOM\n(for\norder\nquantity)", "Unit\nPrice","Season","Buyer\nRequest\nDate","adidas\nCRD","Seller Confirm\nDelivery Date"
            //N
            ,"Color\nMatching"//,"Shipping\nInstruction",
        };

        //public static List<string> RemarkHeaders = new List<string>
        //{
        //    "Remarks","Additional\nOptional 1", "Additional\nOptional 2", "Additional\nOptional 3", "Additional\nOptional 4", "Additional\nOptional 5"
        //};
        public static List<string> RemarkHeaders = new List<string>
        {
            "Remarks"//,"Additional\nOptional 1", "Additional\nOptional 2", "Additional\nOptional 3", "Additional\nOptional 4", "Additional\nOptional 5"
        };

        /// <summary>
        /// Fuhsun xls header
        /// </summary>
        public static string[] FuhsunHeaders = {
            //A                                                   //E
            "T2CODE", "Buyer"/*"VENDOR"*/, "DELIVERYDATE", "FTY", "STYLENO", "STYLEDESCRIPTION", "BSTY0R", "MAINPONO", "Order Number"/*"PURN0R"*/,
            //J                                                      //N
            "VERSION", "STOCKCODE", "ITEMDESCRIPTION", "STOCKCOLOR", "COLORDESCRIPTION", "附件代號", "附件顏色", "尺碼", "數量",
            //S                                                //X
            "單位", "價錢", "貨幣", "季度/SEASON", "季度/YEAR", "MIN", "GMT DATE", "訂單生產排序", "PI DATE", "PI DATE", "REMARK",
            //AD                                                //AI
            "配面布", "W/R數量", "LC0190 Date", "聯絡人", "嘜頭", "SDP", "字軌", "標頭", "備註", "備註", "REMARK-8", "", "", "", "出貨地",
            //AS                                                //AY
            "訂單別", "樣品量", "業務", "業助", "", "訂單+R.L.W", "相同訂單號合併"
        };

        /// <summary>
        /// Adidas xlsx header
        /// </summary>
        public static string[] AdidasHeader = {
            //A
            "Issue Date:", "Order Number:", "Latest Update (GMT):", "Version:", "Purpose:", "Buyer (T1 Supplier Name)",
            //G
            "T1 address line 1", "T1 address line 2", "T1 address line 3", "T1 address line 4", "T1 address city",
            //L
            "T1 address province/ state", "T1 address postal code", "T1 address Country (refer to Pick List_Country)",
            //O
            "Contact Person (Buyer)", "Contact No. (Buyer)", "Email (Buyer)", "Ship To (T1 Factory Name)",
            //S
            "T1 address line 1", "T1 address line 2", "T1 address line 3", "T1 address line 4", "T1 address city",
            //X
            "T1 address province/ state", "T1 address postal code", "T1 address Country (refer to Pick List_Country)",
            //AA
            "T1 6 digit adidas Factory Code", "Seller (T2 Supplier Name)", "Actual Manufacturer (T2 Factory Name)",
            //AD
            "T2 address line 1", "T2 address line 2", "T2 address line 3", "T2 address line 4", "T2 address city",
            //AI
            "T2 address province/ state", "T2 address postal code", "T2 address Country (refer to Pick List_Country)",
            //AL
            "T2 6 digit adidas Factory Code", "T1 Customer", "PO Download Date", "Currency", "Payment Terms",
            //AQ
            "Incoterm", "Ship Mode", "Country of Origin", "Forwarder", "Remarks", "Packing Instruction:",
            //AW
            "Shipping Instruction:", "Line Status", "Line #", "Ref#", "Description / Supplier Material Name", "Width",
            //BE                                                                           //BH      //BI  20171228 add those 2
            "UOM (for width)", "Weight", "UOM (for weight)", "Length", "UOM (for length)", "Height", "UOM (for Height)",
            //BJ
            "Thickness","UOM (for thickness)", "Size", "Material Color", "Unit Price", "Order Quantity", "UOM (for order quantity)",
            //BQ                              //BS 20171228 add
            "Color Matching", "Section Name", "Sustainable material", "Buyer Request Date", "adidas CRD", "adidas Plan Date",
            //BW
            "Seller Confirm Delivery Date", "Seller Updated Delivery Date", "Confirmed Delivery Quantity",
            //BZ
            "Updated Delivery Quantity", "Seller Confirm Delivery Date (Last Shipment)",
            //CB
            "Seller Updated Delivery Date(Last Shipment)", "Confirmed Delivery Quantity(Last Shipment)",
            //CD
            "Updated Delivery Quantity(Last Shipment)", "Season", "Priority & Order Type", "adidas Order Number",
            //CH
            "adidas Article Number", "adidas Working Number / Model Name", "Remarks", "Additional Optional 1",
            //CL
            "Additional Optional 2", "Additional Optional 3", "Additional Optional 4", "Additional Optional 5"
        };

        /// <summary>
        /// Maping from adidas to fuhsun
        /// </summary>
        public static Dictionary<string, string> AdidasToFuhsunMapping = new Dictionary<string, string>
        {
            {"A","AL"},{"B","F"},{"C","BW"},{"D","R"},{"E","CI"},{"I","B"},{"J","D"},{"K","AZ"},{"L","BA"},{"M","BM"}/*N*/,
            {"R","BO"},{"S","BP"},{"T","BN"},{"U","AO"}/*V W X*/,{"Y","BU"},{"AA","BT"},{"AB","BW"},{"AC","CJ"},{"AF","AN"},{"AG","O"},
            {"AL","CK"},{"AM","CL"},{"AN","CM"},{"AO","CN"},{"AP","CO"}
        };

        /// <summary>
        /// Maping form adidas to summary
        /// </summary>
        public static Dictionary<string, string> AdidasToSummaryMapping = new Dictionary<string, string>
        {
            {"A","R"},{"B","CI"},{"C","B"},{"D","AZ"},{"E","BM"},/*F*/
            { "G","BO"},{ "H","BP"},{"I","BN"},{"J","CE"},{"K","BT"},{"L","BU"},{"M","BW"},{"N","BQ"},
            { "O","AW"},{ "P","CJ"},{ "Q","CK"},{"R","CL"},{"S","CM"},{"T","CN"},{"U","CO"}
        };

        /// <summary>
        /// Columns for Summary sheet
        /// </summary>
        public static List<string> SummaryMapping = new List<string>
        {
            //"R","CF","B","AZ","BK","BM","BN","BL","CB","BQ","BR","BT","BO"//,"AW"
            "R","CI","B","AZ","BM","","BO","BP","BN","CE","BT","BU","BW","BQ"
        };

        /// <summary>
        /// Columes for Summary sheet remark part
        /// </summary>
        //public static List<string> RemarkMapping = new List<string>
        //{
        //    "CG","CH","CI","CJ","CK","CL"
        //};

        public static List<string> RemarkMapping = new List<string>
        {
            "CJ"//,"CH","CI","CJ","CK","CL" //
        };

        /// <summary>
        /// set column with datatime
        /// </summary>
        ///                                                                                  BT  BU  BX  BY
        public static HashSet<int> DateTimeColume = new HashSet<int> { 1, 3, 40, 72, 73, 74, 75, 76, 79, 80 };
    }

    public class InnerResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public int Code { get; set; }
        public object Data { get; set; }
    }

    public class Column
    {
        public string Name { get; set; }
        public string Field { get; set; }
        public bool DateTimeColumn { get; set; }
    }


}
