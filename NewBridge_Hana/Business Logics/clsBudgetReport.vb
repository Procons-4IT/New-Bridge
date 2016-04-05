Public Class clsBudgetReport
    Inherits clsBase

    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Dim oStatic As SAPbouiCOM.StaticText
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oCombobox1 As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oBudgetGrid As SAPbouiCOM.Grid
    Private ocombo As SAPbouiCOM.ComboBoxColumn
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Z_BUD_R, frm_Z_BUD_R)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            initialize(oForm)
            loadCombo(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_BUD_R Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 1) Then
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                    changeLabel(oForm)
                                ElseIf pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    LoadBudget(oForm)
                                    oBudgetGrid = oForm.Items.Item("9").Specific
                                    If oBudgetGrid.DataTable.Rows.Count >= 1 Then
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        changeLabel(oForm)
                                    Else
                                        If oBudgetGrid.DataTable.Rows.Count = 0 Then
                                            oApplication.Utilities.Message("No Budget Defined for the Selection...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                ElseIf pVal.ItemUID = "4" Then
                                    oForm.Freeze(True)
                                    If oForm.PaneLevel <> 2 Then
                                        oForm.PaneLevel = 2
                                    Else
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                        changeLabel(oForm)
                                    End If
                                    oForm.Freeze(False)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    reDrawForm(oForm)
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
                Case mnu_Z_BUD_R
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Data Event"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Validations"

    Private Function Validation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strYear, strCategory As String
            
            strYear = CType(oForm.Items.Item("7").Specific, SAPbouiCOM.ComboBox).Value
            strCategory = CType(oForm.Items.Item("8").Specific, SAPbouiCOM.ComboBox).Value
            If strYear.Length = 0 And strCategory.Length = 0 Then
                oApplication.Utilities.Message("Select Year & Category...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strYear = "" Then
                oApplication.Utilities.Message("Select Year ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                'ElseIf strCategory = "" Then
                '    oApplication.Utilities.Message("Select Category ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
            End If

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Items.Item("1").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.Items.Item("17").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.DataSources.DataTables.Add("dtBudget")
            oForm.Items.Item("13").TextStyle = 5
            changeLabel(oForm)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub LoadBudget(ByVal aform As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            Dim strqry As String
            Dim strYear, strCategory As String

            strYear = CType(oForm.Items.Item("7").Specific, SAPbouiCOM.ComboBox).Value.Trim()
            strCategory = CType(oForm.Items.Item("8").Specific, SAPbouiCOM.ComboBox).Value.Trim()

            oBudgetGrid = oForm.Items.Item("9").Specific
            oBudgetGrid.DataTable = oForm.DataSources.DataTables.Item("dtBudget")

            strqry = " Select T0.""U_Year"",T2.""Name"" "
            strqry = strqry & " ,T1.""U_DisCode"" ,T1.""U_OcrCode"",""U_OcrCode2"",T1.""U_OcrCode3"",T1.""U_OcrCode4"" "
            strqry = strqry & " ,T1.""U_OcrCode5"",T1.""U_Z_Budget"",T1.""U_Z_PRApprvd"",""U_Z_POApprvd"",T1.""U_Z_GRApprvd"" "
            strqry = strqry & " ,T1.""U_Z_IVApprvd"",""U_Z_ABudget"" "
            strqry = strqry & " From ""@Z_OBUDDF"" T0 JOIN ""@Z_BUDDF1"" T1 On T0.""DocEntry"" = T1.""DocEntry""  "
            strqry = strqry & " JOIN ""OACG"" T2 On T1.""U_Category"" = T2.""AbsId""  "
            strqry = strqry & " Where T0.""U_Year"" = '" & strYear.Trim() & "'"

            If strCategory.Length > 0 Then
                strqry = strqry & " And T1.""U_Category"" = '" & strCategory.Trim() & "'"
            End If

            oBudgetGrid.DataTable.ExecuteQuery(strqry)

            oBudgetGrid.Columns.Item("U_Year").TitleObject.Caption = "Year"
            oBudgetGrid.Columns.Item("U_Year").Editable = False

            oBudgetGrid.Columns.Item("Name").TitleObject.Caption = "Category Name"
            oBudgetGrid.Columns.Item("Name").Editable = False

            oBudgetGrid.Columns.Item("U_DisCode").TitleObject.Caption = "Effective From"
            oBudgetGrid.Columns.Item("U_DisCode").Editable = False
            oBudgetGrid.Columns.Item("U_DisCode").Visible = False

            oBudgetGrid.Columns.Item("U_OcrCode").TitleObject.Caption = "Dimension 1"
            oBudgetGrid.Columns.Item("U_OcrCode").Editable = False

            oBudgetGrid.Columns.Item("U_OcrCode2").TitleObject.Caption = "Dimension 2"
            oBudgetGrid.Columns.Item("U_OcrCode2").Editable = False

            oBudgetGrid.Columns.Item("U_OcrCode3").TitleObject.Caption = "Dimension 3"
            oBudgetGrid.Columns.Item("U_OcrCode3").Editable = False

            oBudgetGrid.Columns.Item("U_OcrCode4").TitleObject.Caption = "Dimension 4"
            oBudgetGrid.Columns.Item("U_OcrCode4").Editable = False

            oBudgetGrid.Columns.Item("U_OcrCode5").TitleObject.Caption = "Dimension 5"
            oBudgetGrid.Columns.Item("U_OcrCode5").Editable = False

            oBudgetGrid.Columns.Item("U_Z_Budget").TitleObject.Caption = "Total Budget"
            oBudgetGrid.Columns.Item("U_Z_Budget").Editable = False
            oBudgetGrid.Columns.Item("U_Z_Budget").RightJustified = True

            oBudgetGrid.Columns.Item("U_Z_PRApprvd").TitleObject.Caption = "Purchase Request"
            oBudgetGrid.Columns.Item("U_Z_PRApprvd").Editable = False
            oBudgetGrid.Columns.Item("U_Z_PRApprvd").RightJustified = True

            oBudgetGrid.Columns.Item("U_Z_POApprvd").TitleObject.Caption = "Purchase Order"
            oBudgetGrid.Columns.Item("U_Z_POApprvd").Editable = False
            oBudgetGrid.Columns.Item("U_Z_POApprvd").RightJustified = True

            oBudgetGrid.Columns.Item("U_Z_GRApprvd").TitleObject.Caption = "GRPO"
            oBudgetGrid.Columns.Item("U_Z_GRApprvd").Editable = False
            oBudgetGrid.Columns.Item("U_Z_GRApprvd").RightJustified = True

            oBudgetGrid.Columns.Item("U_Z_IVApprvd").TitleObject.Caption = "GRPO"
            oBudgetGrid.Columns.Item("U_Z_IVApprvd").Editable = False
            oBudgetGrid.Columns.Item("U_Z_IVApprvd").RightJustified = True

            oBudgetGrid.Columns.Item("U_Z_ABudget").TitleObject.Caption = "Available Budget"
            oBudgetGrid.Columns.Item("U_Z_ABudget").Editable = False
            oBudgetGrid.Columns.Item("U_Z_ABudget").RightJustified = True

            oBudgetGrid.AutoResizeColumns()
            oBudgetGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub loadCombo(ByVal oForm As SAPbouiCOM.Form)
        Try
            oCombobox = oForm.Items.Item("7").Specific
            Dim year As String
            year = DateTime.Now.Year - 5
            With oComboBox
                While year <= Date.Now().Year + 5
                    .ValidValues.Add(year, year)
                    year = year + 1
                End While
            End With

            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select ""AbsId"" As ""Code"",""Name"" As ""Name"" From ""OACG""")
            If Not oRecordSet.EoF Then
                oCombobox1 = oForm.Items.Item("8").Specific
                Dim i As Integer = 0
                With oComboBox1
                    While i <= oRecordSet.RecordCount - 1
                        .ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                        oRecordSet.MoveNext()
                        i += 1
                    End While
                End With
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            oForm.Items.Item("9").Top = oForm.Items.Item("13").Top + oForm.Items.Item("13").Height + 1
            oForm.Items.Item("9").Height = (oForm.Height - 100)
            oForm.Items.Item("9").Width = oForm.Width - 25
            oForm.Freeze(False)
        Catch ex As Exception
            'oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub changeLabel(ByVal oForm As SAPbouiCOM.Form)
        Try
            oStatic = oForm.Items.Item("17").Specific
            oStatic.Caption = "Step " & oForm.PaneLevel & " of 3"
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

End Class