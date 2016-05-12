
Public Class clsBudgetDefinition
    Inherits clsBase

    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oComboBox, oComboBox1 As SAPbouiCOM.ComboBox
    Private oColumn As SAPbouiCOM.Column
    Private count As Integer
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines_1 As SAPbouiCOM.DBDataSource
    Public MatrixId As String
    Public intSelectedMatrixrow As Integer = 0
    Private RowtoDelete As Integer
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private oCombo As SAPbouiCOM.ComboBox

#Region "Initialization"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "Load Form"

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Z_OBUDDF, frm_Z_OBUDDF)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm)
            loadCombo(oForm)
            loadComboColumn(oForm)
            oForm.DataBrowser.BrowseBy = "3"
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#End Region

    Private Sub initialize(ByVal aform As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OBUDDF")
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_BUDDF1")
            oForm.Items.Item("3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Doc Num
            oForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Doc Date
            oForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Year
            oForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False) 'Category
            AddMode(aform)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "Menu Event"

    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Z_OBUDDF
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddRow(oForm)
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        AddMode(oForm)
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        'enableControls(oForm, True)
                    End If
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#End Region

#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OBUDDF Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            If Validation(oForm) = False Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "7" And pVal.ColUID = "V_0" And pVal.CharPressed = 9 Then
                                    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim oObj As New clsDisRule
                                    clsDisRule.SourceFormUID = FormUID
                                    clsDisRule.ItemUID = pVal.ItemUID
                                    clsDisRule.sourceColumID = pVal.ColUID
                                    clsDisRule.sourcerowId = pVal.Row
                                    Dim stvalue As String
                                    stvalue = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", pVal.Row)
                                    stvalue = stvalue & ";" & oApplication.Utilities.getMatrixValues(oMatrix, "V_2", pVal.Row)
                                    stvalue = stvalue & ";" & oApplication.Utilities.getMatrixValues(oMatrix, "V_3", pVal.Row)
                                    stvalue = stvalue & ";" & oApplication.Utilities.getMatrixValues(oMatrix, "V_4", pVal.Row)
                                    stvalue = stvalue & ";" & oApplication.Utilities.getMatrixValues(oMatrix, "V_5", pVal.Row)
                                    oObj.strStaticValue = stvalue 'oApplication.Utilities.getEdittextvalue(oForm, pVal.ColUID)
                                    oApplication.Utilities.LoadForm(xml_DisRule, frm_DisRule)
                                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    oObj.databound(oForm)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                        End Select
                End Select

            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Data Events"

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            If oForm.TypeEx = frm_Z_OBUDDF Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OBUDDF")
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub loadComboColumn(ByVal oForm As SAPbouiCOM.Form)
        Try
            oMatrix = oForm.Items.Item("7").Specific
            If oMatrix.RowCount = 0 Then
                oMatrix.AddRow(1, -1)
            End If
            oCombo = oMatrix.Columns.Item("V_7").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select ""AbsId"" As ""Code"",""Name"" As ""Name"" From ""OACG"""
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value)
                    oRecordSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case "0", "1"
                    oMatrix = aForm.Items.Item("7").Specific
                    oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_BUDDF1")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    End If
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    Try
                        If oEditText.Value <> "" Then
                            oMatrix.AddRow()
                            oMatrix.ClearRowData(oMatrix.RowCount)
                        End If
                    Catch ex As Exception
                        aForm.Freeze(False)
                    End Try
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines_1.Size
                        oDBDataSourceLines_1.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case "1"
                    oMatrix = aForm.Items.Item("7").Specific
                    oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_BUDDF1")
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines_1.Size
                        oDBDataSourceLines_1.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_BUDDF1")
            If Me.MatrixId = "7" Then
                oMatrix = aForm.Items.Item("7").Specific
                Me.RowtoDelete = intSelectedMatrixrow
                oDBDataSourceLines_1.RemoveRecord(Me.RowtoDelete - 1)
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()
                For count = 1 To oDBDataSourceLines_1.Size
                    oDBDataSourceLines_1.SetValue("LineId", count - 1, count)
                Next
            End If
            oMatrix.LoadFromDataSource()
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub loadCombo(ByVal oForm As SAPbouiCOM.Form)
        Try
            oComboBox = oForm.Items.Item("5").Specific
            Dim year As String
            year = DateTime.Now.Year - 5
            With oComboBox
                While year <= Date.Now().Year + 5
                    .ValidValues.Add(year, year)
                    year = year + 1
                End While
            End With


            oMatrix = oForm.Items.Item("7").Specific
            Dim orecordset As SAPbobsCOM.Recordset
            orecordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oMatrix.RowCount = 0 Then
                oMatrix.AddRow(1, -1)
            End If

            orecordset.DoQuery("select ""DimCode"" as ""code"",""DimDesc"" as ""name"" from ""ODIM""")
            If Not orecordset.EoF Then
                Dim oEditText As SAPbouiCOM.Column = oMatrix.Columns.Item("V_1")
                oEditText.TitleObject.Caption = orecordset.Fields.Item("name").Value
                orecordset.MoveNext()
                oEditText = oMatrix.Columns.Item("V_2")
                oEditText.TitleObject.Caption = orecordset.Fields.Item("name").Value
                orecordset.MoveNext()
                oEditText = oMatrix.Columns.Item("V_3")
                oEditText.TitleObject.Caption = orecordset.Fields.Item("name").Value
                orecordset.MoveNext()
                oEditText = oMatrix.Columns.Item("V_4")
                oEditText.TitleObject.Caption = orecordset.Fields.Item("name").Value
                orecordset.MoveNext()
                oEditText = oMatrix.Columns.Item("V_5")
                oEditText.TitleObject.Caption = orecordset.Fields.Item("name").Value
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AddMode(ByVal aForm As SAPbouiCOM.Form)
        Dim strCode As String
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            strCode = oApplication.Utilities.getMaxCode("@Z_OBUDDF", "DocNum")
            oApplication.Utilities.setEdittextvalue(aForm, "3", strCode)
            oForm.Items.Item("3").Enabled = True
            aForm.Items.Item("4").Enabled = True
            aForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oApplication.Utilities.setEdittextvalue(aForm, "4", "T")
            aForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        End If
    End Sub

    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim blnStatus As Boolean = False
        Try
            ' oMatrix = oForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OBUDDF")
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@Z_BUDDF1")

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If oApplication.Utilities.getEdittextvalue(aForm, "4") = "" Then
                oApplication.Utilities.Message("Select Project Code...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim strYear As String = CType(oForm.Items.Item("5").Specific, SAPbouiCOM.ComboBox).Selected.Value
            'Dim strCategory As String = CType(oForm.Items.Item("6").Specific, SAPbouiCOM.ComboBox).Selected.Value

            If strYear = "" Then
                oApplication.Utilities.Message("Select Year ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            'If strCategory = "" Then
            '    oApplication.Utilities.Message("Select Category ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                strQuery = "Select ""DocEntry"" from ""@Z_OBUDDF"" where " & _
                            """U_Year"" = '" & strYear & "' and ""U_Active"" = 'Y' "
                Try
                    oTest.DoQuery(strQuery)
                Catch ex As Exception

                End Try
                If oTest.RecordCount > 0 Then
                    oApplication.Utilities.Message("Budget already defined for the selected year ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

#End Region

End Class
