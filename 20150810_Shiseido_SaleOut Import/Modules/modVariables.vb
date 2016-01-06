Public Module modVariables
    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String
    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public companylocalCurrency As String
    Public frmSourceMatrix As SAPbouiCOM.Matrix

    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Const frm_WAREHOUSES As Integer = 62
    Public Const frm_ITEM_MASTER As Integer = 150
    Public Const frm_INVOICES As Integer = 133
    Public Const frm_GRPO As Integer = 143
    Public Const frm_ORDR As Integer = 139
    Public Const frm_GR_INVENTORY As Integer = 721
    Public Const frm_Project As Integer = 711
    Public Const frm_ProdReceipt As Integer = 65214
    Public Const frm_Delivery As Integer = 140
    Public Const frm_SaleReturn As Integer = 180
    Public Const frm_ARCreditMemo As Integer = 179
    Public Const frm_Customer As Integer = 134
    Public Const frm_OPOR As Integer = 142
    Public Const frm_OPCH As Integer = 141

    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"

    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"

    Public Const frm_ImpWiz As String = "frm_ImpWiz"
    Public Const xml_ImpWiz As String = "frm_ImpWiz.xml"

    Public Const mnu_SellOut As String = "mnu_SellOut"

End Module
