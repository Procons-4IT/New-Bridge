Public Class clsCustCommission
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
            oForm = oApplication.Utilities.LoadForm(xml_CustRebate, frm_CustComDef)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.DataSources.UserDataSources.Add("Mark", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oApplication.Utilities.setUserDatabind(oForm, "4", "Mark")
            oForm.DataSources.UserDataSources.Add("Pro", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oApplication.Utilities.setUserDatabind(oForm, "6", "Pro")
            oForm.DataSources.UserDataSources.Add("Code", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oApplication.Utilities.setUserDatabind(oForm, "7", "Code")
            oForm.Freeze(True)
            AddChooseFromList(oForm)
            BindData(oForm)
            Databind(oForm)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFL = oCFLs.Item("CFL_2")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


            oCFL = oCFLs.Item("CFL_3")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub BindData(ByVal aForm As SAPbouiCOM.Form)

        oEditText = aForm.Items.Item("4").Specific
        oEditText.ChooseFromListUID = "CFL_2"
        oEditText.ChooseFromListAlias = "Formatcode"

        oEditText = aForm.Items.Item("6").Specific
        oEditText.ChooseFromListUID = "CFL_3"
        oEditText.ChooseFromListAlias = "Formatcode"

    End Sub
#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select * from ""@Z_OCRE"""
            oRec.DoQuery(strQuery)
            If oRec.RecordCount > 0 Then
                oApplication.Utilities.setEdittextvalue(aform, "4", oRec.Fields.Item("U_Z_MarkReb").Value)
                oApplication.Utilities.setEdittextvalue(aform, "6", oRec.Fields.Item("U_Z_ProReb").Value)
                oApplication.Utilities.setEdittextvalue(aform, "7", oRec.Fields.Item("Code").Value)
            End If
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region
#Region "AddtoUDT"
    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode As String
        oUserTable = oApplication.Company.UserTables.Item("Z_OCRE")
        strCode = oApplication.Utilities.getEdittextvalue(aform, "7")
        If oUserTable.GetByKey(strCode) Then
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.UserFields.Fields.Item("U_Z_MarkReb").Value = oApplication.Utilities.getEdittextvalue(aform, "4")
            oUserTable.UserFields.Fields.Item("U_Z_ProReb").Value = oApplication.Utilities.getEdittextvalue(aform, "6")
            If oUserTable.Update <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        Else
            strCode = oApplication.Utilities.getMaxCode("@Z_OCRE", "Code")
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.UserFields.Fields.Item("U_Z_MarkReb").Value = oApplication.Utilities.getEdittextvalue(aform, "4")
            oUserTable.UserFields.Fields.Item("U_Z_ProReb").Value = oApplication.Utilities.getEdittextvalue(aform, "6")
            If oUserTable.Add <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Databind(aform)
        Return True
    End Function

#End Region
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            If oApplication.Utilities.getEdittextvalue(aForm, "4") = "" Then
                oApplication.Utilities.Message("Marketing Rebate Account is missing", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aForm, "6") = "" Then
                oApplication.Utilities.Message("Commission Rebate Account is missing", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_CustComDef Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then
                                    If validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then
                                    AddtoUDT1(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val, val2, val3, val4 As String
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

                                        If pVal.ItemUID = "4" Then
                                            val = oDataTable.GetValue("FormatCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "4", val)
                                        End If
                                        If pVal.ItemUID = "6" Then
                                            val = oDataTable.GetValue("FormatCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "6", val)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    ' oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                Case mnu_CustRebate
                    LoadForm()
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
