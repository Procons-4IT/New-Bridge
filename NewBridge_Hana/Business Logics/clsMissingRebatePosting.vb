Public Class clsMissingRebatePosting
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, strQuery As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Posting, frm_Posting)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.DataSources.UserDataSources.Add("Choice", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oCombobox = oForm.Items.Item("4").Specific
            oCombobox.DataBind.SetBound(True, "", "Choice")
            oCombobox.ValidValues.Add("I", "Invoice")
            oCombobox.ValidValues.Add("C", "Credit Note")
            oCombobox.ValidValues.Add("P", "AP Invoice")
            oCombobox.Select("I", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            oForm.Freeze(True)
         
            'Databind(oForm)
          
            strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date"" ,'Y' ""Select"" from OINV T0 where 1=2"
            oGrid = oForm.Items.Item("9").Specific
            oGrid.DataTable.ExecuteQuery(strSQL)
            oGrid.AutoResizeColumns()
            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Posting Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "8" Then
                                    DataBind(oForm)
                                End If
                                If pVal.ItemUID = "3" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to create rebate posting ?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    Posting(oForm)
                                End If
                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Posting
                    If pVal.BeforeAction - False Then
                        LoadForm()
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region
    Private Sub Posting(aForm As SAPbouiCOM.Form)
        Try

            aForm.Freeze(True)

            Dim aKey As Integer
            Dim strChoice As String
            Dim oCheckbox As SAPbouiCOM.CheckBoxColumn
            oCombobox = aForm.Items.Item("4").Specific
            strChoice = oCombobox.Selected.Value
            oGrid = aForm.Items.Item("9").Specific
            '  oCheckbox = oGrid.Columns.Item("Select")
            oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If strChoice = "I" Then 'Invoice
                Dim oobj As SAPbobsCOM.Documents
                oobj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If 1 = 1 Then ' oCheckbox.IsChecked(intRow) = True Then
                        aKey = oGrid.DataTable.GetValue("InternalKey", intRow)
                        If oobj.GetByKey(aKey) Then
                            If oobj.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items Then
                                If oobj.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                                    CreditAPCreditNote(oobj.DocEntry)
                                    CancelJournal(oobj.DocEntry)
                                Else
                                    CreateAPInvoice(oobj.DocEntry)
                                    CreateJournal(oobj.DocEntry)
                                End If
                            End If
                        End If
                    End If
                Next
            End If

            If strChoice = "P" Then 'AP Invoice
                Dim oobj As SAPbobsCOM.Documents
                oobj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If 1 = 1 Then ' oCheckbox.IsChecked(intRow) = True Then
                        aKey = oGrid.DataTable.GetValue("InternalKey", intRow)
                        If oobj.GetByKey(aKey) Then
                            If oobj.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items Then
                                 If oobj.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                                    CancelAPInvoice_Purchase(oobj.DocEntry)
                                Else
                                    CreditAPCreditNote_APInvoice(oobj.DocEntry)
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            If strChoice = "C" Then 'AR Credit Note
                Dim oobj As SAPbobsCOM.Documents
                oobj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If 1 = 1 Then 'oCheckbox.IsChecked(intRow) = True Then
                        aKey = oGrid.DataTable.GetValue("InternalKey", intRow)
                        If oobj.GetByKey(aKey) Then
                            If oobj.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items Then

                                If oobj.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                                    CancelARCreditNoe(oobj.DocEntry)
                                Else
                                    CreditAPCreditNote_ARCreditNote(oobj.DocEntry)
                                End If
                            End If
                        End If

                    End If
                Next

            End If
            oApplication.Utilities.Message("Operation Completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Function CancelAPInvoice_Purchase(ByVal DocNum As String) As Boolean
        Dim oAPInv, oAPInv2 As SAPbobsCOM.Documents
        Dim oAPInv1 As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        oAPInv2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Try
            '  ApplyRebateAmount(DocNum)
            strQuery = "Select * from OPCH where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                If oTest.Fields.Item("U_Z_BaseEntry").Value <> "" Then
                    oRec.DoQuery("Select * from ORPC where ""DocEntry""=" & oTest.Fields.Item("U_Z_BaseEntry").Value)
                    If oRec.RecordCount <= 0 Then
                        Return True
                    Else
                        '  MsgBox(oRec.Fields.Item("DocEntry").Value)
                    End If
                    If oAPInv2.GetByKey(oRec.Fields.Item("DocEntry").Value) Then
                        If oAPInv2.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oAPInv1 = oAPInv2.CreateCancellationDocument()
                            If oAPInv1.Add() <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
    Public Function CreditAPCreditNote_APInvoice(ByVal DocNum As String) As Boolean
        Dim oAPInv, oAPInv2 As SAPbobsCOM.Documents
        Dim oAPInv1 As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        oAPInv2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Try
            '  ApplyRebateAmount(DocNum)
            strQuery = "Select * from OPCH where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                If 1 = 2 Then 'oTest.Fields.Item("U_Z_BaseEntry").Value <> "" Then
                    oRec.DoQuery("Select * from ORPC where ""DocEntry""=" & oTest.Fields.Item("U_Z_BaseEntry").Value)
                    If oRec.RecordCount <= 0 Then
                        Return True
                    Else
                        '  MsgBox(oRec.Fields.Item("DocEntry").Value)
                    End If
                    If oAPInv.GetByKey(oRec.Fields.Item("DocEntry").Value) Then
                        If oAPInv.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oAPInv2 = oAPInv.CreateCancellationDocument()
                            If oAPInv2.Add() <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        Else
                            oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                            Dim oWhs As SAPbobsCOM.Warehouses
                            Dim OItem As SAPbobsCOM.Items
                            oWhs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWarehouses)
                            OItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            oAPInv1.DocDate = oTest.Fields.Item("DocDate").Value
                            oAPInv1.DocDueDate = oTest.Fields.Item("DocDueDate").Value
                            oAPInv1.CardCode = oAPInv.CardCode
                            oAPInv1.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                            oAPInv1.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                            oAPInv1.Comments = "Supplier Rebate Posting - Based on A/P Invoice  : " & oTest.Fields.Item("DocNum").Value.ToString
                            oAPInv1.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                            Dim intLineCount As Integer = 0
                            Dim dblUnitPrice, dblPercentage As Double
                            Dim blnlineExist As Boolean = False
                            For intLoop As Integer = 0 To oAPInv.Lines.Count - 1
                                oAPInv.Lines.SetCurrentLine(intLoop)
                                If oWhs.GetByKey(oAPInv.Lines.WarehouseCode) Then
                                    If oWhs.UserFields.Fields.Item("U_Z_Type").Value = "D" And oWhs.UserFields.Fields.Item("U_Z_PWhs").Value <> "" And oAPInv.Lines.UserFields.Fields.Item("U_Z_Rebate").Value > 0 Then
                                        If intLineCount > 0 Then
                                            oAPInv1.Lines.Add()
                                            oAPInv1.Lines.SetCurrentLine(intLineCount)
                                        End If
                                        oAPInv1.Lines.ItemCode = oAPInv.Lines.AccountCode
                                        oAPInv1.Lines.ItemDescription = oAPInv.Lines.ItemDescription
                                        dblPercentage = oAPInv.Lines.UserFields.Fields.Item("U_Z_RebValue").Value

                                        oAPInv1.Lines.Currency = oAPInv.Lines.Currency
                                        oAPInv1.Lines.UnitPrice = oAPInv.Lines.UnitPrice
                                        oAPInv1.Lines.Quantity = oAPInv.Lines.Quantity
                                        oAPInv1.Lines.WithoutInventoryMovement = SAPbobsCOM.BoYesNoEnum.tYES
                                        oAPInv1.Lines.WarehouseCode = oWhs.UserFields.Fields.Item("U_Z_PWhs").Value
                                        If oAPInv.Lines.RowTotalFC > 0 Then
                                            dblUnitPrice = oAPInv.Lines.RowTotalFC
                                            dblUnitPrice = (dblUnitPrice) - (dblUnitPrice * dblPercentage / 100)
                                            oAPInv1.Lines.RowTotalFC = dblUnitPrice
                                        Else
                                            dblUnitPrice = oAPInv.Lines.LineTotal
                                            dblUnitPrice = (dblUnitPrice) - (dblUnitPrice * dblPercentage / 100)
                                            oAPInv1.Lines.LineTotal = dblUnitPrice
                                        End If

                                        'oAPInv1.Lines.BaseType = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                                        'oAPInv1.Lines.BaseEntry = oAPInv.DocEntry
                                        'oAPInv1.Lines.BaseLine = oAPInv.Lines.LineNum
                                        blnlineExist = True
                                        intLineCount = intLineCount + 1
                                    End If
                                End If

                            Next
                            If blnlineExist = True Then
                                If oAPInv1.Add <> 0 Then
                                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Return False
                                Else
                                    oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim strDocNum As String
                                    oApplication.Company.GetNewObjectCode(strDocNum)
                                    oAPInv1.GetByKey(CInt(strDocNum))
                                    strDocNum = oAPInv1.DocNum
                                    strQuery = "Update OPCH set ""U_Z_BaseEntry""='" & oAPInv1.DocEntry.ToString & "', ""U_Z_APInvoice""='" & strDocNum & "' where ""DocEntry""=" & DocNum
                                    oTest2.DoQuery(strQuery)
                                End If
                            End If
                        End If
                    End If
                Else
                    If 1 = 1 Then ' oAPInv.GetByKey(oRec.Fields.Item("DocEntry").Value) Then
                        If 1 = 2 Then ' oAPInv.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oAPInv2 = oAPInv.CreateCancellationDocument()
                            If oAPInv2.Add() <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        Else
                            oAPInv.GetByKey(DocNum)
                            oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                            Dim oWhs As SAPbobsCOM.Warehouses
                            Dim OItem As SAPbobsCOM.Items
                            oWhs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWarehouses)
                            OItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            oAPInv1.DocDate = oAPInv.DocDate
                            oAPInv1.DocDueDate = oAPInv.DocDueDate
                            oAPInv1.CardCode = oAPInv.CardCode
                            oAPInv1.UserFields.Fields.Item("U_Z_ARInvoice").Value = oAPInv.DocNum.ToString
                            oAPInv1.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                            oAPInv1.Comments = "Supplier Rebate Posting - Based on A/P Invoice  : " & oAPInv.DocNum.ToString
                            oAPInv1.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                            Dim intLineCount As Integer = 0
                            Dim dblUnitPrice, dblPercentage As Double
                            Dim blnlineExist As Boolean = False
                            For intLoop As Integer = 0 To oAPInv.Lines.Count - 1
                                oAPInv.Lines.SetCurrentLine(intLoop)
                                ' MsgBox(oAPInv.Lines.WarehouseCode)
                                If oWhs.GetByKey(oAPInv.Lines.WarehouseCode) Then
                                    ' If oWhs.UserFields.Fields.Item("U_Z_Type").Value = "D" And oWhs.UserFields.Fields.Item("U_Z_PWhs").Value <> "" And oAPInv.Lines.UserFields.Fields.Item("U_Z_Rebate").Value > 0 Then
                                    If oWhs.UserFields.Fields.Item("U_Z_Type").Value = "D" And oAPInv.Lines.UserFields.Fields.Item("U_Z_Rebate").Value > 0 Then
                                        If intLineCount > 0 Then
                                            oAPInv1.Lines.Add()
                                            oAPInv1.Lines.SetCurrentLine(intLineCount)
                                        End If
                                        oAPInv1.Lines.ItemCode = oAPInv.Lines.ItemCode
                                        oAPInv1.Lines.ItemDescription = oAPInv.Lines.ItemDescription
                                        dblUnitPrice = oAPInv.Lines.UnitPrice
                                        dblPercentage = oAPInv.Lines.UserFields.Fields.Item("U_Z_Rebate").Value
                                        oAPInv1.Lines.Currency = oAPInv.Lines.Currency
                                        '     oAPInv1.Lines.UnitPrice = oAPInv.Lines.UnitPrice
                                        oAPInv1.Lines.Quantity = oAPInv.Lines.Quantity
                                        oAPInv1.Lines.WithoutInventoryMovement = SAPbobsCOM.BoYesNoEnum.tYES
                                        oAPInv1.Lines.WarehouseCode = oAPInv.Lines.WarehouseCode
                                        If oAPInv.Lines.RowTotalFC > 0 Then
                                            dblUnitPrice = oAPInv.Lines.RowTotalFC
                                            dblUnitPrice = (dblUnitPrice) - (dblUnitPrice * dblPercentage / 100)
                                            oAPInv1.Lines.RowTotalFC = dblUnitPrice
                                            oAPInv1.Lines.UnitPrice = dblUnitPrice / oAPInv.Lines.Quantity
                                        Else
                                            dblUnitPrice = oAPInv.Lines.LineTotal
                                            dblUnitPrice = (dblUnitPrice) - (dblUnitPrice * dblPercentage / 100)
                                            oAPInv1.Lines.LineTotal = dblUnitPrice
                                            oAPInv1.Lines.UnitPrice = dblUnitPrice / oAPInv.Lines.Quantity
                                        End If
                                        oAPInv1.Lines.BaseType = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                                        oAPInv1.Lines.BaseEntry = oAPInv.DocEntry
                                        oAPInv1.Lines.BaseLine = oAPInv.Lines.LineNum
                                        blnlineExist = True
                                        intLineCount = intLineCount + 1
                                    End If
                                End If
                            Next
                            If blnlineExist = True Then
                                If oAPInv1.Add <> 0 Then
                                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Return False
                                Else
                                    oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim strDocNum As String
                                    oApplication.Company.GetNewObjectCode(strDocNum)
                                    oAPInv1.GetByKey(CInt(strDocNum))
                                    strDocNum = oAPInv1.DocNum
                                    strQuery = "Update OPCH set ""U_Z_BaseEntry""='" & oAPInv1.DocEntry.ToString & "', ""U_Z_APInvoice""='" & strDocNum & "' where ""DocEntry""=" & DocNum
                                    oTest2.DoQuery(strQuery)
                                End If
                            End If
                        End If
                    End If

                End If
                Return True
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_Invoice Then
                    Dim oobj As SAPbobsCOM.Documents
                    oobj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    If oobj.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oobj.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                            CreditAPCreditNote(oobj.DocEntry)
                            CancelJournal(oobj.DocEntry)
                        Else
                            CreateAPInvoice(oobj.DocEntry)
                            CreateJournal(oobj.DocEntry)
                        End If
                    End If
                End If
                If oForm.TypeEx = frm_ARCreditNote Then
                    Dim oobj As SAPbobsCOM.Documents
                    oobj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    If oobj.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oobj.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                            '  CreditAPCreditNote_ARCreditNote(oobj.DocEntry)
                            CancelARCreditNoe(oobj.DocEntry)
                        Else
                            CreditAPCreditNote_ARCreditNote(oobj.DocEntry)
                        End If
                    End If

                End If
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub ApplyRebateAmount(ByVal aDocEntry As Integer)
        Dim oAPInv As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        Dim strCardCode, strItemCode, strQuery, strDistRule As String
        Dim dblLineTotal, dblCommission, dblMarketing, dblComPercentage, dblMarketingPercentage, dblSupComm, dblSupCommPreLicense, dblItemCost, dblUnitPrice, dblQty, dblHubTotal As Double
        Dim dtPostingDate As Date
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        Dim blnHubWhs As Boolean = False
        ' oRec.DoQuery("Select T1.""CardCode"" ""Card"", * from INV1 T0 Inner Join OINV T1 on T1.""DocEntry""=T0.""DocEntry"" inner Join OITM T2 on T2.""ItemCode""=T0.""ItemCode"" where T0.""DocEntry""=" & aDocEntry)
        oRec.DoQuery("Select T0.""BaseCard"" ""Card"", * from INV1 T0 Inner Join OINV T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocEntry""=" & aDocEntry)
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strCardCode = oRec.Fields.Item("Card").Value
            strItemCode = oRec.Fields.Item("ItemCode").Value
            dtPostingDate = oRec.Fields.Item("DocDate").Value
            dblUnitPrice = oRec.Fields.Item("PriceBefDi").Value
            dblQty = oRec.Fields.Item("Quantity").Value

            oTemp1.DoQuery("Select * from OITM where ""ItemCode""='" & strItemCode & "'")
            dblSupComm = oTemp1.Fields.Item("U_Z_SupCom").Value
            dblSupCommPreLicense = oTemp1.Fields.Item("U_Z_SupComPre").Value


            otest1.DoQuery("Select ifnull(""U_Z_Type"",'R') from OWHS where ""WhsCode""='" & oRec.Fields.Item("WhsCode").Value & "'")
            If otest1.Fields.Item(0).Value = "H" Then
                blnHubWhs = True
                otest1.DoQuery("Select ""AvgPrice"" from OITW where ""WhsCode""='" & oRec.Fields.Item("WhsCode").Value & "' and ""ItemCode""='" & strItemCode & "'")
                'If otest1.RecordCount > 0 Then
                '    dblItemCost = otest1.Fields.Item(0).Value
                'Else
                '    dblItemCost = 0
                'End If
                dblItemCost = oRec.Fields.Item("StockPrice").Value
            Else
                dblItemCost = 0

            End If

            strDistRule = oRec.Fields.Item("OcrCode").Value ' & ";" & oRec.Fields.Item("OcrCode2").Value & ";" & oRec.Fields.Item("OcrCode3").Value & ";" & oRec.Fields.Item("OcrCode4").Value & ";" & oRec.Fields.Item("OcrCode5").Value
            strQuery = "Select T0.""ItemCode"",T1.""U_Z_Comm"",T1.""U_Z_MarkReb"",T1.""U_Z_RegStatus"" From OSPP T0 LEFT OUTER JOIN SPP1 T1 On T0.""ItemCode"" = T1.""ItemCode"" And T0.""CardCode"" = T1.""CardCode""  Where T1.""U_Z_OcrCode""='" & strDistRule & "' and T0.""CardCode"" = '" & strCardCode & "' And T0.""ItemCode"" = '" & strItemCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between  ifnull(T1.""FromDate"",NOW()) and ifnull(T1.""ToDate"",NOW())"
            otest1.DoQuery(strQuery)
            If otest1.RecordCount > 0 Then
                Dim oTe As SAPbobsCOM.Recordset
                oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTe.DoQuery("Select ifnull(""U_Z_ComBase"",0) from OCRD where  ""CardCode""='" & strCardCode & "'")
                If oTe.Fields.Item(0).Value > 0 Then
                    dblUnitPrice = (dblUnitPrice * oTe.Fields.Item(0).Value / 100)
                    dblLineTotal = dblUnitPrice * dblQty
                Else
                    dblLineTotal = oRec.Fields.Item("LineTotal").Value
                End If
                '   dblLineTotal = oRec.Fields.Item("LineTotal").Value
                dblCommission = otest1.Fields.Item(1).Value
                dblMarketing = otest1.Fields.Item(2).Value
                dblComPercentage = dblCommission
                dblMarketingPercentage = dblMarketing
                If dblCommission <> 0 Then
                    dblCommission = dblLineTotal * dblCommission / 100
                Else
                    dblCommission = 0
                    dblComPercentage = 0
                End If
                If dblMarketing <> 0 Then
                    dblMarketing = dblLineTotal * dblMarketing / 100
                Else
                    dblMarketing = 0
                    dblMarketingPercentage = 0
                End If
                If blnHubWhs = True Then
                    If otest1.Fields.Item("U_Z_RegStatus").Value = "R" Or otest1.Fields.Item("U_Z_RegStatus").Value = "N" Then
                        Dim s As String = "Update INV1 set ""U_Z_ItemCost""='" & dblItemCost & "', ""U_Z_SupCom1""='" & dblSupComm & "', ""U_Z_RegStatus""='" & otest1.Fields.Item("U_Z_RegStatus").Value & "', ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value
                        dblHubTotal = (dblUnitPrice * dblQty * dblSupComm / 100)
                        dblHubTotal = dblHubTotal - (dblQty * dblItemCost)

                        oTemp1.DoQuery("Update INV1 set ""U_Z_Accrual""='" & dblHubTotal & "', ""U_Z_ItemCost""='" & dblItemCost & "', ""U_Z_SupCom1""='" & dblSupComm & "', ""U_Z_RegStatus""='" & otest1.Fields.Item("U_Z_RegStatus").Value & "', ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
                    ElseIf otest1.Fields.Item("U_Z_RegStatus").Value = "L" Then
                        dblHubTotal = (dblUnitPrice * dblQty * dblSupCommPreLicense / 100)
                        dblHubTotal = dblHubTotal - (dblQty * dblItemCost)
                        oTemp1.DoQuery("Update INV1 set ""U_Z_Accrual""='" & dblHubTotal & "',  ""U_Z_ItemCost""='" & dblItemCost & "', ""U_Z_SupCom1""='" & dblSupCommPreLicense & "', ""U_Z_RegStatus""='" & otest1.Fields.Item("U_Z_RegStatus").Value & "', ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
                    Else
                        oTemp1.DoQuery("Update INV1 set  ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
                    End If
                Else
                    oTemp1.DoQuery("Update INV1 set  ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & aDocEntry & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
                End If
            Else
                dblCommission = 0
                dblMarketing = 0
                oTemp1.DoQuery("Update INV1 set  ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "',""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='N' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
            End If
            oRec.MoveNext()
        Next
    End Sub

    Public Sub ApplyRebateAmount_CreditNote(ByVal aDocEntry As Integer)
        Dim oAPInv As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        Dim strCardCode, strItemCode, strQuery, strDistRule As String
        Dim dblLineTotal, dblCommission, dblMarketing, dblComPercentage, dblMarketingPercentage As Double
        Dim dtPostingDate As Date
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        oRec.DoQuery("Select T0.""BaseCard"" ""Card"", * from RIN1 T0 Inner Join ORIN T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocEntry""=" & aDocEntry)
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strCardCode = oRec.Fields.Item("Card").Value
            strItemCode = oRec.Fields.Item("ItemCode").Value
            dtPostingDate = oRec.Fields.Item("DocDate").Value
            strDistRule = oRec.Fields.Item("OcrCode").Value ' & ";" & oRec.Fields.Item("OcrCode2").Value & ";" & oRec.Fields.Item("OcrCode3").Value & ";" & oRec.Fields.Item("OcrCode4").Value & ";" & oRec.Fields.Item("OcrCode5").Value
            'strQuery = "Select T0.""ItemCode"",T1.""U_Z_Comm"",T1.""U_Z_MarkReb"" From OSPP T0 LEFT OUTER JOIN SPP1 T1 On T0.""ItemCode"" = T1.""ItemCode"" And T0.""CardCode"" = T1.""CardCode""  Where T1.""U_Z_OcrCode""='" & strDistRule & "' and T0.""CardCode"" = '" & strCardCode & "' And T0.""ItemCode"" = '" & strItemCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between  ifnull(T1.""FromDate"",NOW()) and ifnull(T1.""ToDate"",NOW())"
            strQuery = "Select T0.""ItemCode"",T1.""U_Z_Comm"",T1.""U_Z_MarkReb"" From OSPP T0 LEFT OUTER JOIN SPP1 T1 On T0.""ItemCode"" = T1.""ItemCode"" And T0.""CardCode"" = T1.""CardCode""  Where T1.""U_Z_OcrCode""='" & strDistRule & "' and T0.""CardCode"" = '" & strCardCode & "' And T0.""ItemCode"" = '" & strItemCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between  ifnull(T1.""FromDate"",NOW()) and ifnull(T1.""ToDate"",NOW())"
            'otest1.DoQuery(strQuery)
            If oRec.Fields.Item("U_Z_IsComm").Value = "Y" And (oRec.Fields.Item("U_Z_Comm_Per").Value <> 0 Or oRec.Fields.Item("U_Z_MarkReb_Per").Value <> 0) Then
                dblLineTotal = oRec.Fields.Item("LineTotal").Value
                dblCommission = oRec.Fields.Item("U_Z_Comm_Per").Value
                dblMarketing = oRec.Fields.Item("U_Z_MarkReb_Per").Value
                dblComPercentage = dblCommission
                dblMarketingPercentage = dblMarketing
                If dblCommission <> 0 Then
                    dblCommission = dblLineTotal * dblCommission / 100
                Else
                    dblCommission = 0
                    dblComPercentage = 0
                End If
                If dblMarketing <> 0 Then
                    dblMarketing = dblLineTotal * dblMarketing / 100
                Else
                    dblMarketing = 0
                    dblMarketingPercentage = 0
                End If
                otest1.DoQuery("Update RIN1 set ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
            Else
                dblCommission = 0
                dblMarketing = 0
                otest1.DoQuery("Update RIN1 set ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "',""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='N' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
            End If
            oRec.MoveNext()
        Next
    End Sub

    Public Function CreateJournal(ByVal DocNum As String) As Boolean
        Dim oAPInv As SAPbobsCOM.JournalEntries
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        Dim strCreditAc, strDebitAc As String

        Try
            strCreditAc = ""
            strDebitAc = ""


            If 1 = 1 Then 'strCreditAc <> "" And strDebitAc <> "" Then
                '   strQuery = "Select * from OINV where ""DocEntry""=" & DocNum & ""
                '  oTest.DoQuery(strQuery) '

                If oApplication.Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    strQuery = "Select * from OINV where ifnull(""U_Z_JournalRef"",'')='' and ""DocEntry""=" & DocNum & ""
                Else
                    strQuery = "Select * from OINV where isnull(""U_Z_JournalRef"",'')='' and ""DocEntry""=" & DocNum & ""
                End If
                If oTest.RecordCount > 0 Then
                    strQuery = "Select * from OCRD where ""CardCode""='" & oTest.Fields.Item("CardCode").Value & "' and ""U_Z_ComRePay""<>''"
                    otest1.DoQuery(strQuery)
                    If otest1.RecordCount > 0 Then
                        strQuery = "Select * from OCRD where ""CardCode""='" & otest1.Fields.Item("U_Z_ComRePay").Value & "' and ""CardType""='S'"
                        oRec.DoQuery(strQuery)
                        If oRec.RecordCount > 0 Then
                            oAPInv.TaxDate = oTest.Fields.Item("DocDate").Value
                            oAPInv.DueDate = oTest.Fields.Item("DocDueDate").Value
                            oAPInv.ReferenceDate = oTest.Fields.Item("DocDate").Value
                            'oAPInv.CardCode = otest1.Fields.Item("U_Z_ComRePay").Value
                            oAPInv.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                            oAPInv.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                            oAPInv.Memo = "Rebate Posting Based on A/R Invoice  : " & oTest.Fields.Item("DocNum").Value.ToString
                            Dim blnLineExists As Boolean = False
                            strQuery = "Select (""U_Z_Accrual"") ""U_Z_Comm"",""OcrCode"",""OcrCode2"",""OcrCode3"",""OcrCode4"",""OcrCode5"",T1.""ItmsGrpCod"" from INV1 T0 Inner Join OITM T1 on T1.""ItemCode""=T0.""ItemCode"" where T0.""DocEntry""='" & DocNum & "' and ""U_Z_RegStatus""<>'O'"
                            oTemp1.DoQuery(strQuery)
                            Dim dbLComm, dblMarketing As Double
                            Dim dblDebit As Double = 0
                            Dim strCountry As String = ""
                            Dim intLineCount As Integer = 0
                            For intloop As Integer = 0 To oTemp1.RecordCount - 1
                                dbLComm = 0
                                dblMarketing = 0
                                dbLComm = oTemp1.Fields.Item("U_Z_Comm").Value
                                If dbLComm <> 0 Then
                                    oTest.DoQuery("Select * from ""@Z_OSRE"" where ""U_Z_ItmsGrp""='" & oTemp1.Fields.Item("ItmsGrpCod").Value & "'")
                                    If oTest.RecordCount > 0 Then
                                        strDebitAc = oTest.Fields.Item("U_Z_COGS").Value
                                        strCreditAc = oTest.Fields.Item("U_Z_Accrual").Value
                                    Else
                                        strDebitAc = ""
                                        strCreditAc = ""
                                    End If
                                    If strDebitAc <> "" And strDebitAc <> "" Then
                                        'Credit Entry
                                        If intLineCount > 0 Then
                                            oAPInv.Lines.Add()
                                        End If
                                        oAPInv.Lines.SetCurrentLine(intLineCount)
                                        dblDebit = dblDebit + dbLComm
                                        oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(strCreditAc)
                                        oAPInv.Lines.Credit = dbLComm
                                        oAPInv.Lines.Reference2 = otest1.Fields.Item("U_Z_ComRePay").Value
                                        If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                            oAPInv.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                        End If
                                        strCountry = oTemp1.Fields.Item("OcrCode").Value
                                        If oTemp1.Fields.Item("OcrCode2").Value <> "" Then
                                            oAPInv.Lines.CostingCode2 = oTemp1.Fields.Item("OcrCode2").Value
                                        End If
                                        If oTemp1.Fields.Item("OcrCode3").Value <> "" Then
                                            oAPInv.Lines.CostingCode3 = oTemp1.Fields.Item("OcrCode3").Value
                                        End If
                                        If oTemp1.Fields.Item("OcrCode4").Value <> "" Then
                                            oAPInv.Lines.CostingCode4 = oTemp1.Fields.Item("OcrCode4").Value
                                        End If
                                        If oTemp1.Fields.Item("OcrCode5").Value <> "" Then
                                            oAPInv.Lines.CostingCode5 = oTemp1.Fields.Item("OcrCode5").Value
                                        End If
                                        intLineCount = intLineCount + 1
                                        blnLineExists = True
                                        'Debit Entry
                                        If intLineCount > 0 Then
                                            oAPInv.Lines.Add()
                                        End If
                                        oAPInv.Lines.SetCurrentLine(intLineCount)
                                        oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(strDebitAc)
                                        oAPInv.Lines.Debit = dbLComm
                                        oAPInv.Lines.Reference2 = otest1.Fields.Item("U_Z_ComRePay").Value
                                        If strCountry <> "" Then
                                            oAPInv.Lines.CostingCode = strCountry
                                        End If
                                        If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                            oAPInv.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                        End If
                                        strCountry = oTemp1.Fields.Item("OcrCode").Value
                                        If oTemp1.Fields.Item("OcrCode2").Value <> "" Then
                                            oAPInv.Lines.CostingCode2 = oTemp1.Fields.Item("OcrCode2").Value
                                        End If
                                        If oTemp1.Fields.Item("OcrCode3").Value <> "" Then
                                            oAPInv.Lines.CostingCode3 = oTemp1.Fields.Item("OcrCode3").Value
                                        End If
                                        If oTemp1.Fields.Item("OcrCode4").Value <> "" Then
                                            oAPInv.Lines.CostingCode4 = oTemp1.Fields.Item("OcrCode4").Value
                                        End If
                                        If oTemp1.Fields.Item("OcrCode5").Value <> "" Then
                                            oAPInv.Lines.CostingCode5 = oTemp1.Fields.Item("OcrCode5").Value
                                        End If
                                        intLineCount = intLineCount + 1
                                        blnLineExists = True
                                    End If
                                End If
                                oTemp1.MoveNext()
                            Next
                            If blnLineExists = True Then
                                If oAPInv.Add <> 0 Then
                                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Return False
                                Else
                                    oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim strDocNum As String
                                    oApplication.Company.GetNewObjectCode(strDocNum)
                                    oAPInv.GetByKey(CInt(strDocNum))
                                    strDocNum = oAPInv.JdtNum
                                    strQuery = "Update OINV set ""U_Z_JournalRef""='" & strDocNum & "' where ""DocEntry""=" & DocNum
                                    oTest2.DoQuery(strQuery)

                                End If
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Private Sub DataBind(aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strSQL, strFromDate, strToDate, strChoice, strcondition As String
            Dim dtFromDate, dtToDate As Date
            oCombobox = aform.Items.Item("4").Specific
            strChoice = oCombobox.Selected.Value
            strFromDate = oApplication.Utilities.getEdittextvalue(aform, "6")
            strToDate = oApplication.Utilities.getEdittextvalue(aform, "7")
            If strFromDate = "" Then
                strcondition = "1=1"
            Else
                dtFromDate = oApplication.Utilities.GetDateTimeValue(strFromDate)
                strcondition = " T0.""DocDate"" >='" & dtFromDate.ToString("yyyy-MM-dd") & "'"
            End If

            If strToDate = "" Then
                strcondition = strcondition & " and 1=1"
            Else
                dtToDate = oApplication.Utilities.GetDateTimeValue(strToDate)
                strcondition = strcondition & " and  T0.""DocDate"" <='" & dtFromDate.ToString("yyyy-MM-dd") & "'"
            End If
            If oApplication.Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                If strChoice = "I" Then
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date"" ,'Y' ""Select"" from OINV T0 where T0.""DocType""='I' and " & strcondition & " and ifnull(T0.""U_Z_APInvoice"",'')=''"
                ElseIf strChoice = "P" Then
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date"" ,'Y' ""Select"" from OPCH T0 where T0.""DocType""='I' and " & strcondition & " and ifnull(T0.""U_Z_APInvoice"",'')=''"

                Else
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date"" ,'Y' ""Select"" from ORIN T0 where  T0.""DocType""='I' and " & strcondition & " and ifnull(T0.""U_Z_APInvoice"",'')=''"

                End If
            Else
                If strChoice = "I" Then
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date"" ,'Y' ""Select"" from OINV T0 where T0.""DocType""='I' and " & strcondition & " and isnull(T0.""U_Z_APInvoice"",'')=''"
                ElseIf strChoice = "P" Then
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date"" ,'Y' ""Select"" from OPCH T0 where T0.""DocType""='I' and " & strcondition & " and isnull(T0.""U_Z_APInvoice"",'')=''"

                Else
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date"" ,'Y' ""Select"" from ORIN T0 where T0.""DocType""='I' and " & strcondition & " and isnull(T0.""U_Z_APInvoice"",'')=''"

                End If
            End If



            If oApplication.Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                If strChoice = "I" Then
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date""  from OINV T0 where T0.""DocType""='I' and " & strcondition & " and ifnull(T0.""U_Z_APInvoice"",'')=''"
                ElseIf strChoice = "P" Then
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date""   from OPCH T0 where T0.""DocType""='I' and " & strcondition & " and ifnull(T0.""U_Z_APInvoice"",'')=''"
                Else
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date""  from ORIN T0 where  T0.""DocType""='I' and " & strcondition & " and ifnull(T0.""U_Z_APInvoice"",'')=''"
                End If
            Else
                If strChoice = "I" Then
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date""  from OINV T0 where T0.""DocType""='I' and " & strcondition & " and isnull(T0.""U_Z_APInvoice"",'')=''"
                ElseIf strChoice = "P" Then
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date""   from OPCH T0 where T0.""DocType""='I' and " & strcondition & " and isnull(T0.""U_Z_APInvoice"",'')=''"
                Else
                    strSQL = "Select T0.""DocEntry"" ""InternalKey"",T0.""DocNum"" ""Document Number"",T0.""CardCode"" ""Customer Code"", T0.""CardName"" ""Customer Name"",T0.""DocDate"" ""Document Date""  from ORIN T0 where T0.""DocType""='I' and " & strcondition & " and isnull(T0.""U_Z_APInvoice"",'')=''"
                End If
            End If
           
            oGrid = aform.Items.Item("9").Specific
            oGrid.DataTable.ExecuteQuery(strSQL)
            oEditTextColumn = oGrid.Columns.Item(2)
            oEditTextColumn.LinkedObjectType = "2"

            oEditTextColumn = oGrid.Columns.Item(0)
            If strChoice = "I" Then
                oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice
            ElseIf strChoice = "P" Then
                oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_PurchaseInvoice
            Else

                oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo
            End If

            ' oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

            aform.Freeze(False)
        Catch ex As Exception
            aform.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Function CreateAPInvoice(ByVal DocNum As String) As Boolean
        Dim oAPInv As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        Try

            If oApplication.Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                strQuery = "Select * from OINV where ifnull(""U_Z_APInvoice"",'')='' and ""DocEntry""=" & DocNum & ""
            Else
                strQuery = "Select * from OINV where isnull(""U_Z_APInvoice"",'')='' and ""DocEntry""=" & DocNum & ""
            End If
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                ApplyRebateAmount(DocNum)
                strQuery = "Select * from OCRD where ""CardCode""='" & oTest.Fields.Item("CardCode").Value & "' and ""U_Z_ComRePay""<>''"
                otest1.DoQuery(strQuery)
                If otest1.RecordCount > 0 Then
                    strQuery = "Select * from OCRD where ""CardCode""='" & otest1.Fields.Item("U_Z_ComRePay").Value & "' and ""CardType""='S'"
                    oRec.DoQuery(strQuery)
                    If oRec.RecordCount > 0 Then
                        oAPInv.DocDate = oTest.Fields.Item("DocDate").Value
                        oAPInv.DocDueDate = oTest.Fields.Item("DocDueDate").Value
                        oAPInv.CardCode = otest1.Fields.Item("U_Z_ComRePay").Value
                        oAPInv.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                        oAPInv.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                        oAPInv.NumAtCard = "AR Invoice No : " & oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv.Comments = "Rebate/Commission posting based on A/R Invoice No  : " & oTest.Fields.Item("DocNum").Value.ToString
                        Dim blnLineExists As Boolean = False
                        strQuery = "Select (""U_Z_Comm"") ""U_Z_Comm"",(""U_Z_MarkReb"") ""U_Z_MarkReb"",""OcrCode"",""OcrCode2"",""OcrCode3"",""OcrCode4"",""OcrCode5"" from INV1 where ""DocEntry""='" & DocNum & "' and ""U_Z_IsComm""='Y' "
                        oTemp1.DoQuery(strQuery)
                        Dim dbLComm, dblMarketing As Double
                        Dim intLineCount As Integer = 0
                        For intloop As Integer = 0 To oTemp1.RecordCount - 1
                            dbLComm = 0
                            dblMarketing = 0
                            dbLComm = oTemp1.Fields.Item("U_Z_Comm").Value
                            dblMarketing = oTemp1.Fields.Item("U_Z_MarkReb").Value
                            If intLineCount > 0 Then
                                oAPInv.Lines.Add()
                            End If
                            oAPInv.Lines.SetCurrentLine(intLineCount)
                            oTest.DoQuery("Select * from ""@Z_OCRE""")
                            If dbLComm > 0 Then
                                oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTest.Fields.Item("U_Z_ProReb").Value) 'oTest.Fields.Item("U_Z_MarkReb").Value
                                intLineCount = intLineCount + 1
                                oAPInv.Lines.LineTotal = dbLComm
                                oAPInv.Lines.ItemDescription = "Commission Rebate"
                                blnLineExists = True
                                If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                    oAPInv.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode2").Value <> "" Then
                                    oAPInv.Lines.CostingCode2 = oTemp1.Fields.Item("OcrCode2").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode3").Value <> "" Then
                                    oAPInv.Lines.CostingCode3 = oTemp1.Fields.Item("OcrCode3").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode4").Value <> "" Then
                                    oAPInv.Lines.CostingCode4 = oTemp1.Fields.Item("OcrCode4").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode5").Value <> "" Then
                                    oAPInv.Lines.CostingCode5 = oTemp1.Fields.Item("OcrCode5").Value
                                End If
                            End If
                            If intLineCount > 0 Then
                                oAPInv.Lines.Add()
                            End If
                            oAPInv.Lines.SetCurrentLine(intLineCount)
                            If dblMarketing > 0 Then
                                oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTest.Fields.Item("U_Z_MarkReb").Value) 'oTest.Fields.Item("U_Z_MarkReb").Value
                                intLineCount = intLineCount + 1
                                oAPInv.Lines.LineTotal = dblMarketing
                                oAPInv.Lines.ItemDescription = "Marketing  Rebate"
                                blnLineExists = True
                                If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                    oAPInv.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode2").Value <> "" Then
                                    oAPInv.Lines.CostingCode2 = oTemp1.Fields.Item("OcrCode2").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode3").Value <> "" Then
                                    oAPInv.Lines.CostingCode3 = oTemp1.Fields.Item("OcrCode3").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode4").Value <> "" Then
                                    oAPInv.Lines.CostingCode4 = oTemp1.Fields.Item("OcrCode4").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode5").Value <> "" Then
                                    oAPInv.Lines.CostingCode5 = oTemp1.Fields.Item("OcrCode5").Value
                                End If
                            End If
                            oTemp1.MoveNext()
                        Next
                        If blnLineExists = True Then
                            If oAPInv.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            Else
                                oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                oAPInv.GetByKey(CInt(strDocNum))
                                strDocNum = oAPInv.DocNum
                                strQuery = "Update OINV set ""U_Z_BaseEntry""='" & oAPInv.DocEntry.ToString & "', ""U_Z_APInvoice""='" & strDocNum & "' where ""DocEntry""=" & DocNum
                                oTest2.DoQuery(strQuery)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
    Public Function CreditAPCreditNote(ByVal DocNum As String) As Boolean
        Dim oAPInv, oAPInv2 As SAPbobsCOM.Documents
        Dim oAPInv1 As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        oAPInv2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Try
            '  ApplyRebateAmount(DocNum)
            strQuery = "Select * from OINV where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                If oTest.Fields.Item("U_Z_APInvoice").Value <> "" Then
                    oRec.DoQuery("Select * from OPCH where ""DocNum""=" & oTest.Fields.Item("U_Z_APInvoice").Value)
                    If oRec.RecordCount <= 0 Then
                        Return True
                    Else
                        '  MsgBox(oRec.Fields.Item("DocEntry").Value)
                    End If

                    If oAPInv.GetByKey(oRec.Fields.Item("DocEntry").Value) Then
                        If oAPInv.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oAPInv2 = oAPInv.CreateCancellationDocument()
                            If oAPInv2.Add() <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        Else
                            oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                            oAPInv1.DocDate = oTest.Fields.Item("DocDate").Value
                            oAPInv1.DocDueDate = oTest.Fields.Item("DocDueDate").Value
                            oAPInv1.CardCode = oAPInv.CardCode
                            oAPInv1.NumAtCard = oAPInv.NumAtCard
                            oAPInv1.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                            oAPInv1.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                            oAPInv1.Comments = "Rebate/Commission posting canceled based on A/P Invoice  : " & oTest.Fields.Item("DocNum").Value.ToString
                            oAPInv1.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                            For intLoop As Integer = 0 To oAPInv.Lines.Count - 1
                                If intLoop > 0 Then
                                    oAPInv1.Lines.Add()
                                    oAPInv1.Lines.SetCurrentLine(intLoop)
                                End If
                                oAPInv.Lines.SetCurrentLine(intLoop)
                                oAPInv1.Lines.AccountCode = oAPInv.Lines.AccountCode
                                oAPInv1.Lines.ItemDescription = oAPInv.Lines.ItemDescription
                                oAPInv1.Lines.LineTotal = oAPInv.Lines.LineTotal
                            Next
                            If oAPInv1.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            Else
                                oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                oAPInv1.GetByKey(CInt(strDocNum))
                                strDocNum = oAPInv1.DocNum
                                strQuery = "Update OINV set ""U_Z_BaseEntry""='" & oAPInv1.DocEntry.ToString & "', ""U_Z_APInvoice""='" & strDocNum & "' where ""DocEntry""=" & DocNum
                                oTest2.DoQuery(strQuery)
                            End If
                        End If
                    End If

                End If
                Return True
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function CancelJournal(ByVal DocNum As String) As Boolean
        Dim oAPInv, oAPInv2 As SAPbobsCOM.JournalEntries
        Dim oAPInv1 As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        oAPInv2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Try
            '  ApplyRebateAmount(DocNum)
            strQuery = "Select * from OINV where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                If oTest.Fields.Item("U_Z_JournalRef").Value <> "" Then
                    oRec.DoQuery("Select * from OJDT where ""TransId""=" & oTest.Fields.Item("U_Z_JournalRef").Value)
                    If oRec.RecordCount <= 0 Then
                        Return True
                    Else
                        '  MsgBox(oRec.Fields.Item("DocEntry").Value)
                    End If

                    If oAPInv.GetByKey(oRec.Fields.Item("TransId").Value) Then
                        If 1 = 1 Then 'oAPInv.DocumentStatu = SAPbobsCOM.BoStatus.bost_Open Then
                            If oAPInv.Cancel <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If

                    End If

                End If

            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function CancelARCreditNoe(ByVal DocNum As String) As Boolean
        Dim oAPInv, oAPInv2 As SAPbobsCOM.Documents
        Dim oAPInv1 As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        oAPInv2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Try
            '  ApplyRebateAmount(DocNum)
            strQuery = "Select * from ORIN where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                If oTest.Fields.Item("U_Z_APInvoice").Value <> "" Then
                    oRec.DoQuery("Select * from ORPC where ""DocNum""=" & oTest.Fields.Item("U_Z_APInvoice").Value)
                    If oRec.RecordCount <= 0 Then
                        Return True
                    Else
                        '  MsgBox(oRec.Fields.Item("DocEntry").Value)
                    End If

                    If oAPInv.GetByKey(oRec.Fields.Item("DocEntry").Value) Then
                        If oAPInv.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oAPInv2 = oAPInv.CreateCancellationDocument()
                            If oAPInv2.Add() <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        Else
                            oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                            oAPInv1.DocDate = oTest.Fields.Item("DocDate").Value
                            oAPInv1.DocDueDate = oTest.Fields.Item("DocDueDate").Value
                            oAPInv1.CardCode = oAPInv.CardCode
                            oAPInv1.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                            oAPInv1.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                            oAPInv1.NumAtCard = oAPInv.NumAtCard
                            oAPInv1.Comments = "Rebate/Commission posting -canceled Based on A/R Credit Note  : " & oTest.Fields.Item("DocNum").Value.ToString
                            oAPInv1.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                            For intLoop As Integer = 0 To oAPInv.Lines.Count - 1
                                If intLoop > 0 Then
                                    oAPInv1.Lines.Add()
                                    oAPInv1.Lines.SetCurrentLine(intLoop)
                                End If
                                oAPInv.Lines.SetCurrentLine(intLoop)
                                oAPInv1.Lines.AccountCode = oAPInv.Lines.AccountCode
                                oAPInv1.Lines.ItemDescription = oAPInv.Lines.ItemDescription
                                oAPInv1.Lines.LineTotal = oAPInv.Lines.LineTotal
                            Next
                            If oAPInv1.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            Else
                                oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                oAPInv1.GetByKey(CInt(strDocNum))
                                strDocNum = oAPInv1.DocNum
                                strQuery = "Update OINV set ""U_Z_BaseEntry""='" & oAPInv1.DocEntry.ToString & "', ""U_Z_APInvoice""='" & strDocNum & "' where ""Docentry""=" & DocNum
                                oTest2.DoQuery(strQuery)
                            End If
                        End If
                    End If

                End If
                Return True
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
    Public Function CreditAPCreditNote_ARCreditNote(ByVal DocNum As String) As Boolean
        Dim oAPInv, oAPInv2 As SAPbobsCOM.Documents
        Dim oAPInv1 As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        oAPInv2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Try
            ApplyRebateAmount_CreditNote(DocNum)
            strQuery = "Select * from ORIN where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                If 1 = 1 Then ' oAPInv.GetByKey(oTest.Fields.Item("DocEntry").Value) Then
                    strQuery = "Select * from OCRD where ""CardCode""='" & oTest.Fields.Item("CardCode").Value & "' and ""U_Z_ComRePay""<>''"
                    otest1.DoQuery(strQuery)
                    If otest1.RecordCount <= 0 Then 'oAPInv.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                        'oAPInv2 = oAPInv.CreateCancellationDocument()
                        'If oAPInv2.Add() <> 0 Then
                        '    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False
                        'End If
                    Else
                        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                        oAPInv1.DocDate = oTest.Fields.Item("DocDate").Value
                        oAPInv1.DocDueDate = oTest.Fields.Item("DocDueDate").Value
                        oAPInv1.CardCode = otest1.Fields.Item("U_Z_ComRePay").Value
                        oAPInv1.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv1.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                        oAPInv1.NumAtCard = "AR Credit Memo No : " & oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv1.Comments = "Rebate/Commission posting based on A/R Credit Memo  : " & oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv1.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                        strQuery = "Select sum(""U_Z_Comm"") ""U_Z_Comm"",sum(""U_Z_MarkReb"") ""U_Z_MarkReb"",""OcrCode"" from RIN1 where ""DocEntry""='" & DocNum & "' and ""U_Z_IsComm""='Y' group by ""OcrCode"""
                        oTemp1.DoQuery(strQuery)
                        Dim dbLComm, dblMarketing As Double
                        Dim blnLineExists As Boolean = False
                        Dim intLineCount As Integer = 0
                        For intloop As Integer = 0 To oTemp1.RecordCount - 1
                            dbLComm = 0
                            dblMarketing = 0
                            dbLComm = oTemp1.Fields.Item("U_Z_Comm").Value
                            dblMarketing = oTemp1.Fields.Item("U_Z_MarkReb").Value

                            oTest.DoQuery("Select * from ""@Z_OCRE""")
                            If dbLComm > 0 Then
                                If intLineCount > 0 Then
                                    oAPInv1.Lines.Add()
                                End If
                                oAPInv1.Lines.SetCurrentLine(intLineCount)
                                oAPInv1.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTest.Fields.Item("U_Z_ProReb").Value) 'oTest.Fields.Item("U_Z_MarkReb").Value
                                oAPInv1.Lines.LineTotal = dbLComm
                                oAPInv1.Lines.ItemDescription = "Commission Rebate"
                                blnLineExists = True
                                If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                    oAPInv1.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                End If
                                intLineCount = intLineCount + 1
                            End If
                            If dblMarketing > 0 Then
                                If intLineCount > 0 Then
                                    oAPInv1.Lines.Add()
                                End If
                                '  oAPInv.Lines.SetCurrentLine(intLineCount)
                                oAPInv1.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTest.Fields.Item("U_Z_MarkReb").Value) 'oTest.Fields.Item("U_Z_MarkReb").Value

                                oAPInv1.Lines.LineTotal = dblMarketing
                                oAPInv1.Lines.ItemDescription = "Marketing  Rebate"
                                blnLineExists = True
                                If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                    oAPInv1.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                End If
                                intLineCount = intLineCount + 1
                            End If
                            oTemp1.MoveNext()
                        Next
                        If blnLineExists = True Then
                            If oAPInv1.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            Else
                                oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                oAPInv1.GetByKey(CInt(strDocNum))
                                strDocNum = oAPInv1.DocNum
                                strQuery = "Update ORIN set ""U_Z_BaseEntry""='" & oAPInv1.DocEntry.ToString & "', ""U_Z_APInvoice""='" & strDocNum & "' where ""DocEntry""=" & DocNum
                                oTest2.DoQuery(strQuery)
                            End If
                        End If
                    End If

                End If
                Return True


            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
End Class
