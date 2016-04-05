Public Module modVariables

    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String
    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public strCardCode As String = ""
    Public blnDraft As Boolean = False
    Public blnError As Boolean = False
    Public frmSourceForm As SAPbouiCOM.Form
    Public strDocEntry As String
    Public blnIsHana As Boolean

    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocumentType As Integer
        RENTAL_QUOTATION = 1
        RENTAL_ORDER
        RENTAL_RETURN
    End Enum

    Public Const frm_WAREHOUSES As Integer = 62

    Public Const frm_StockRequest As String = "frm_StRequest"
    Public Const frm_InvSO As String = "frm_InvSO"
    Public Const frm_Warehouse As String = "62"
    Public Const frm_SalesOrder As String = "139"
    Public Const frm_Invoice As String = "133"
    Public Const frm_APInvoice As String = "141"
    Public Const frm_ARCreditNote As String = "179"
  
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
    Public Const mnu_CloseOrderLines As String = "DABT_910"
    Public Const mnu_InvSO As String = "DABT_911"
    
    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"
    Public Const xml_StRequest As String = "StRequest.xml"
    Public Const xml_InvSO As String = "frm_InvSO.xml"

    Public Const frm_SalePriceBP As String = "333"


    Public Const frm_CustComDef As String = "frm_CustComDef"
    Public Const xml_CustRebate As String = "frm_CustComDef.xml"
    Public Const mnu_CustRebate As String = "Mnu_CustRebate"

    Public Const frm_SubComDef As String = "frm_SubComDef"
    Public Const xml_SubRebate As String = "frm_SubComDef.xml"
    Public Const mnu_SubRebate As String = "Mnu_SubRebate"

    Public Const frm_DisRule As String = "frm_DisRule"
    Public Const xml_DisRule As String = "frm_DisRule.xml"

    Public Const frm_LoginSetup As String = "frm_LogSetup"
    Public Const mnu_Logsetup As String = "mnu_003"
    Public Const xml_Logsetup As String = "frm_LoginSetup.xml"

    Public Const frm_Expenses As String = "frm_Expenses"
    Public Const mnu_Expenses As String = "mnu_004"
    Public Const xml_Expenses As String = "frm_Expenses.xml"

    Public Const frm_S01 As String = "frm_S01"
    Public Const mnu_AppTemp As String = "mnu_005"
    Public Const xml_AppTemp As String = "frm_ApprovalTemp.xml"

    Public Const frm_Posting As String = "frm_Posting"
    Public Const mnu_Posting As String = "mnu_006"
    Public Const xml_Posting As String = "frm_Posting.xml"

    Public Const frm_Z_OBUDDF As String = "frm_Z_OBUDDF"
    Public Const mnu_Z_OBUDDF As String = "mnu_007"
    Public Const xml_Z_OBUDDF As String = "frm_Z_OBUDDF.xml"

    Public Const frm_Z_BUD_R As String = "frm_Z_BUD_R"
    Public Const mnu_Z_BUD_R As String = "mnu_008"
    Public Const xml_Z_BUD_R As String = "frm_Z_BUD_R.xml"

End Module
