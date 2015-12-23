Public Class clsSupplierRebate
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

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_APInvoice Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "3" Then
                                    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim strItem As String = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                    Dim oTest As SAPbobsCOM.Items
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                    oTest.GetByKey(strItem)
                                    Try
                                        If oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Rebate", pVal.Row)) <= 0 Then
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Rebate", pVal.Row, oTest.UserFields.Fields.Item("U_Z_RebValue").Value.ToString)
                                        End If
                                       
                                    Catch ex As Exception

                                    End Try



                                End If


                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        '  val = oDataTable.GetValue("OcrCode", 0)
                                        ' oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try

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
                Case mnu_InvSO
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                If oForm.TypeEx = frm_APInvoice Then
                    Dim oobj As SAPbobsCOM.Documents
                    oobj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                    If oobj.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oobj.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                            CancelAPInvoice(oobj.DocEntry)
                        Else
                            CreditAPCreditNote(oobj.DocEntry)
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
        Dim dblLineTotal, dblCommission, dblMarketing, dblComPercentage, dblMarketingPercentage As Double
        Dim dtPostingDate As Date
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        oRec.DoQuery("Select * from INV1 T0 Inner Join OINV T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocEntry""=" & aDocEntry)
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strCardCode = oRec.Fields.Item("CardCode").Value
            strItemCode = oRec.Fields.Item("ItemCode").Value
            dtPostingDate = oRec.Fields.Item("DocDate").Value
            strDistRule = oRec.Fields.Item("OcrCode").Value ' & ";" & oRec.Fields.Item("OcrCode2").Value & ";" & oRec.Fields.Item("OcrCode3").Value & ";" & oRec.Fields.Item("OcrCode4").Value & ";" & oRec.Fields.Item("OcrCode5").Value
            strQuery = "Select T0.""ItemCode"",T1.""U_Z_Comm"",T1.""U_Z_MarkReb"" From OSPP T0 LEFT OUTER JOIN SPP1 T1 On T0.""ItemCode"" = T1.""ItemCode"" And T0.""CardCode"" = T1.""CardCode""  Where T1.""U_Z_OcrCode""='" & strDistRule & "' and T0.""CardCode"" = '" & strCardCode & "' And T0.""ItemCode"" = '" & strItemCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between  ifnull(T1.""FromDate"",NOW()) and ifnull(T1.""ToDate"",NOW())"
            otest1.DoQuery(strQuery)
            If otest1.RecordCount > 0 Then
                dblLineTotal = oRec.Fields.Item("LineTotal").Value
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

                otest1.DoQuery("Update INV1 set ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
            Else
                dblCommission = 0
                dblMarketing = 0
                otest1.DoQuery("Update INV1 set ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "',""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='N' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
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
        oRec.DoQuery("Select * from RIN1 T0 Inner Join ORIN T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocEntry""=" & aDocEntry)
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strCardCode = oRec.Fields.Item("CardCode").Value
            strItemCode = oRec.Fields.Item("ItemCode").Value
            dtPostingDate = oRec.Fields.Item("DocDate").Value
            strDistRule = oRec.Fields.Item("OcrCode").Value ' & ";" & oRec.Fields.Item("OcrCode2").Value & ";" & oRec.Fields.Item("OcrCode3").Value & ";" & oRec.Fields.Item("OcrCode4").Value & ";" & oRec.Fields.Item("OcrCode5").Value
            strQuery = "Select T0.""ItemCode"",T1.""U_Z_Comm"",T1.""U_Z_MarkReb"" From OSPP T0 LEFT OUTER JOIN SPP1 T1 On T0.""ItemCode"" = T1.""ItemCode"" And T0.""CardCode"" = T1.""CardCode""  Where T1.""U_Z_OcrCode""='" & strDistRule & "' and T0.""CardCode"" = '" & strCardCode & "' And T0.""ItemCode"" = '" & strItemCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between  ifnull(T1.""FromDate"",NOW()) and ifnull(T1.""ToDate"",NOW())"
            'otest1.DoQuery(strQuery)
            If oRec.Fields.Item("U_Z_Comm_Per").Value <> 0 Or oRec.Fields.Item("U_Z_MarkReb_Per").Value <> 0 Then
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


    Public Function CreateAPInvoice(ByVal DocNum As String) As Boolean
        Dim oAPInv As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        Try
            ApplyRebateAmount(DocNum)
            strQuery = "Select * from OINV where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
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
                        oAPInv.Comments = "Rebate Posting Based on A/R Invoice  : " & oTest.Fields.Item("DocNum").Value.ToString
                        Dim blnLineExists As Boolean = False
                        '   strQuery = "Select sum(""U_Z_Comm"") 'U_Z_Comm',sum(""U_Z_MarkReb"") 'U_Z_MarkReb',""OcrCode"",""OcrCode2"",""OcrCode3"",""OcrCode4"",""OcrCode5"" from INV1 where ""DocEntry""='" & DocNum & "' and ""U_Z_IsComm""='Y' group by ""OcrCode"",""OcrCode2"",""OcrCode3"",""OcrCode4"",""OcrCode5"""
                        strQuery = "Select sum(""U_Z_Comm"") ""U_Z_Comm"",sum(""U_Z_MarkReb"") ""U_Z_MarkRe"",""OcrCode"" from INV1 where ""DocEntry""='" & DocNum & "' and ""U_Z_IsComm""='Y' group by ""OcrCode"""
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

    Public Function CancelAPInvoice(ByVal DocNum As String) As Boolean
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
                            oAPInv1.Comments = "Rebate Posting -Canceled Based on A/R Credit Note  : " & oTest.Fields.Item("DocNum").Value.ToString
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
                        oAPInv1.Comments = "Rebate Posting - Based on A/R Credit Memo  : " & oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv1.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                        strQuery = "Select Sum(""U_Z_Comm"") ""U_Z_Comm"",sum(""U_Z_MarkReb"") ""U_Z_MarkReb"",""OcrCode"" from RIN1 where ""DocEntry""='" & DocNum & "' and ""U_Z_IsComm""='Y' group by ""OcrCode"""
                        oTemp1.DoQuery(strQuery)
                        Dim dbLComm, dblMarketing As Double
                        Dim blnLineExists As Boolean = False
                        Dim intLineCount As Integer = 0
                        For intloop As Integer = 0 To oTemp1.RecordCount - 1
                            dbLComm = 0
                            dblMarketing = 0
                            dbLComm = oTemp1.Fields.Item("U_Z_Comm").Value
                            dblMarketing = oTemp1.Fields.Item("U_Z_MarkReb").Value
                            If intLineCount > 0 Then
                                oAPInv.Lines.Add()
                            End If
                            oAPInv1.Lines.SetCurrentLine(intLineCount)
                            oTest.DoQuery("Select * from ""@Z_OCRE""")
                            If dbLComm > 0 Then
                                oAPInv1.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTest.Fields.Item("U_Z_ProReb").Value) 'oTest.Fields.Item("U_Z_MarkReb").Value
                                intLineCount = intLineCount + 1
                                oAPInv1.Lines.LineTotal = dbLComm
                                oAPInv1.Lines.ItemDescription = "Commission Rebate"
                                blnLineExists = True
                                If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                    oAPInv.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                End If
                            End If
                            If intLineCount > 0 Then
                                oAPInv.Lines.Add()
                            End If
                            oAPInv.Lines.SetCurrentLine(intLineCount)
                            If dblMarketing > 0 Then
                                oAPInv1.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTest.Fields.Item("U_Z_MarkReb").Value) 'oTest.Fields.Item("U_Z_MarkReb").Value
                                intLineCount = intLineCount + 1
                                oAPInv1.Lines.LineTotal = dblMarketing
                                oAPInv1.Lines.ItemDescription = "Marketing  Rebate"
                                blnLineExists = True
                                If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                    oAPInv.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                End If

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
