Public Class clsSupplierCommission2
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBoxColumn
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
            oForm = oApplication.Utilities.LoadForm(xml_SubRebate, frm_SubComDef)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.DataSources.UserDataSources.Add("Mark", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'oApplication.Utilities.setUserDatabind(oForm, "4", "Mark")
            'oForm.DataSources.UserDataSources.Add("Pro", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'oApplication.Utilities.setUserDatabind(oForm, "6", "Pro")
            'oForm.DataSources.UserDataSources.Add("Code", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'oApplication.Utilities.setUserDatabind(oForm, "7", "Code")
            oForm.Freeze(True)
            oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
            oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
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

        'oEditText = aForm.Items.Item("4").Specific
        'oEditText.ChooseFromListUID = "CFL_2"
        'oEditText.ChooseFromListAlias = "Formatcode"

        'oEditText = aForm.Items.Item("6").Specific
        'oEditText.ChooseFromListUID = "CFL_3"
        'oEditText.ChooseFromListAlias = "Formatcode"

    End Sub
#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select ""Code"",""Name"",""U_Z_ItmsGrp"",""U_Z_COGS"",""U_Z_Accrual"" from ""@Z_OSRE"""
            oGrid = aform.Items.Item("5").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)
            oGrid.Columns.Item("Code").Visible = False
            oGrid.Columns.Item("Name").Visible = False
            oGrid.Columns.Item("U_Z_COGS").TitleObject.Caption = "COGS Account"
            oEditTextColumn = oGrid.Columns.Item("U_Z_COGS")
            oEditTextColumn.ChooseFromListUID = "CFL_2"
            oEditTextColumn.ChooseFromListAlias = "FormatCode"
            oEditTextColumn.LinkedObjectType = "1"
            oGrid.Columns.Item("U_Z_Accrual").TitleObject.Caption = "Accural"
            oEditTextColumn = oGrid.Columns.Item("U_Z_Accrual")
            oEditTextColumn.ChooseFromListUID = "CFL_3"
            oEditTextColumn.ChooseFromListAlias = "FormatCode"
            oEditTextColumn.LinkedObjectType = "1"
            oGrid.Columns.Item("U_Z_ItmsGrp").TitleObject.Caption = "Item Group"
            oGrid.Columns.Item("U_Z_ItmsGrp").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oCombobox = oGrid.Columns.Item("U_Z_ItmsGrp")
            oRec.DoQuery("Select * from OITB")
            For intRow As Integer = 0 To oRec.RecordCount - 1
                oCombobox.ValidValues.Add(oRec.Fields.Item("ItmsGrpCod").Value, oRec.Fields.Item("ItmsGrpNam").Value)
                oRec.MoveNext()
            Next
            oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
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
        Dim strCode, strItmsGrp As String
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_OSRE")
        'strCode = oApplication.Utilities.getEdittextvalue(aform, "7")
        oGrid = aform.Items.Item("5").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            strCode = oGrid.DataTable.GetValue("Code", intRow)
            oCombobox = oGrid.Columns.Item("U_Z_ItmsGrp")
            strItmsGrp = oCombobox.GetSelectedValue(intRow).Value
            If strItmsGrp <> "" Then


                If strCode <> "" Then
                    If oUserTable.GetByKey(strCode) Then
                        oUserTable.Code = strCode
                        oUserTable.Name = strCode
                        oUserTable.UserFields.Fields.Item("U_Z_COGS").Value = oGrid.DataTable.GetValue("U_Z_COGS", intRow)
                        oUserTable.UserFields.Fields.Item("U_Z_Accrual").Value = oGrid.DataTable.GetValue("U_Z_Accrual", intRow)
                        oUserTable.UserFields.Fields.Item("U_Z_ItmsGrp").Value = strItmsGrp
                        If oUserTable.Update <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Else
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_COGS").Value = oGrid.DataTable.GetValue("U_Z_COGS", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_Accrual").Value = oGrid.DataTable.GetValue("U_Z_Accrual", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_ItmsGrp").Value = strItmsGrp
                    If oUserTable.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            End If

        Next
        oRec.DoQuery("Delete from ""@Z_OSRE"" where ""Name"" Like '%_XD'")
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Databind(aform)
        Return True
    End Function

#End Region
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            If oApplication.Utilities.getEdittextvalue(aForm, "4") = "" Then
                oApplication.Utilities.Message("COGS Account is missing", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aForm, "6") = "" Then
                oApplication.Utilities.Message("Accrual Account is missing", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
            If pVal.FormTypeEx = frm_SubComDef Then
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
                                If pVal.ItemUID = "2" Then
                                    Dim oRec As SAPbobsCOM.Recordset
                                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRec.DoQuery("Update ""@Z_OSRE"" set ""Name""=""Code"" ")
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
                                        If oCFL.ObjectType = "1" Then
                                            oGrid = oForm.Items.Item("5").Specific
                                            Try
                                                val = oDataTable.GetValue("FormatCode", 0)
                                                oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                            Catch ex As Exception

                                            End Try
                                        End If
                                        'If pVal.ItemUID = "4" Then
                                        '    val = oDataTable.GetValue("FormatCode", 0)
                                        '    oApplication.Utilities.setEdittextvalue(oForm, "4", val)
                                        'End If
                                        'If pVal.ItemUID = "6" Then
                                        '    val = oDataTable.GetValue("FormatCode", 0)
                                        '    oApplication.Utilities.setEdittextvalue(oForm, "6", val)
                                        'End If
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
                Case mnu_SubRebate
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If oForm.TypeEx = frm_SubComDef Then
                        If pVal.BeforeAction = False Then
                            oGrid = oForm.Items.Item("5").Specific
                            If oGrid.DataTable.GetValue("U_Z_ItmsGrp", oGrid.DataTable.Rows.Count - 1) <> "" Then
                                oGrid.DataTable.Rows.Add()
                            End If
                        End If
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If oForm.TypeEx = frm_SubComDef Then
                        If pVal.BeforeAction = True Then
                            oGrid = oForm.Items.Item("5").Specific
                            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                If oGrid.Rows.IsSelected(intRow) = True Then
                                    Dim oRec As SAPbobsCOM.Recordset
                                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    If oGrid.DataTable.GetValue("Code", intRow) <> "" Then
                                        oRec.DoQuery("Update ""@Z_OSRE"" set ""Name""=Concat(""Name"",'_XD') where ""Code""='" & oGrid.DataTable.GetValue("Code", intRow) & "'")
                                        oGrid.DataTable.Rows.Remove(intRow)
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oGrid.DataTable.Rows.Remove(intRow)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Next
                        End If
                    End If

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
