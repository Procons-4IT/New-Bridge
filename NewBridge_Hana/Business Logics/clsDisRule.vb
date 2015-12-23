Public Class clsDisRule
    Inherits clsBase

#Region "Declarations"
    Public Shared ItemUID As String
    Public Shared SourceFormUID As String
    Public Shared SourceLabel As Integer
    Public Shared CFLChoice As String
    Public Shared ItemCode As String
    Public Shared sourceItemCode As String
    Public Shared choice As String
    Public Shared sourceColumID As String
    Public Shared sourcerowId As Integer
    Public Shared BinDescrUID As String
    Public Shared Documentchoice As String

    Private oDbDatasource As SAPbouiCOM.DBDataSource
    Private Ouserdatasource As SAPbouiCOM.UserDataSource
    Private oConditions As SAPbouiCOM.Conditions
    Private ocondition As SAPbouiCOM.Condition
    Private intRowId As Integer
    Private strRowNum As Integer
    Private i As Integer
    Private oedit As SAPbouiCOM.EditText
    '   Private oForm As SAPbouiCOM.Form
    Private objSoureceForm As SAPbouiCOM.Form
    Private objform As SAPbouiCOM.Form
    Private oMatrix As SAPbouiCOM.Grid
    Private osourcegrid As SAPbouiCOM.Grid
    Private Const SEPRATOR As String = "~~~"
    Private SelectedRow As Integer
    Private sSearchColumn As String
    Private oItem As SAPbouiCOM.Item
    Public stritemid As SAPbouiCOM.Item
    Private intformmode As SAPbouiCOM.BoFormMode
    Private objGrid As SAPbouiCOM.Grid
    Private objSourcematrix As SAPbouiCOM.Matrix
    Private dtTemp As SAPbouiCOM.DataTable
    Private objStatic As SAPbouiCOM.StaticText
    Private inttable As Integer = 0
    Public strformid As String
    Public strStaticValue As String
    Public strSQL As String
    Private strSelectedItem1 As String
    Private strSelectedItem2 As String
    Private strSelectedItem3 As String
    Private strSelectedItem4 As String
    Private oRecSet As SAPbobsCOM.Recordset
    '   Private objSBOAPI As ClsSBO
    '   Dim objTransfer As clsTransfer
#End Region

#Region "New"
    '*****************************************************************
    'Type               : Constructor
    'Name               : New
    'Parameter          : 
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Create object for classes.
    '******************************************************************
    Public Sub New()
        '   objSBOAPI = New ClsSBO
        MyBase.New()
    End Sub
#End Region


#Region "Add Choose From List"
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
            oCon.Alias = "DimCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "1"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL_3")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DimCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "2"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL_4")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DimCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "3"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL_5")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DimCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "4"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()
            oCFL = oCFLs.Item("CFL_6")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DimCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "5"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region
#Region "Bind Data"
    '******************************************************************
    'Type               : Procedure
    'Name               : BindData
    'Parameter          : Form
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Binding the fields.
    '******************************************************************
    Public Sub databound(ByVal objform As SAPbouiCOM.Form)
        Try
            Dim strSQL As String = ""
            Dim ObjSegRecSet As SAPbobsCOM.Recordset
            ObjSegRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objform.Freeze(True)
            objform.DataSources.DataTables.Add("dtLevel3")
            AddChooseFromList(objform)
            Ouserdatasource = objform.DataSources.UserDataSources.Add("dbFind", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            Ouserdatasource = objform.DataSources.UserDataSources.Add("dbFind1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            Ouserdatasource = objform.DataSources.UserDataSources.Add("dbFind2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            Ouserdatasource = objform.DataSources.UserDataSources.Add("dbFind3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            Ouserdatasource = objform.DataSources.UserDataSources.Add("dbFind4", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            Dim ststring As String()
            ststring = strStaticValue.Split(";")

            oedit = objform.Items.Item("ed1").Specific
            oedit.DataBind.SetBound(True, "", "dbFind")
            oedit.ChooseFromListUID = "CFL_2"
            oedit.ChooseFromListAlias = "OcrCode"
            oedit = objform.Items.Item("ed2").Specific
            oedit.DataBind.SetBound(True, "", "dbFind1")
            oedit.ChooseFromListUID = "CFL_3"
            oedit.ChooseFromListAlias = "OcrCode"
            oedit = objform.Items.Item("ed3").Specific
            oedit.DataBind.SetBound(True, "", "dbFind2")
            oedit.ChooseFromListUID = "CFL_4"
            oedit.ChooseFromListAlias = "OcrCode"
            oedit = objform.Items.Item("ed4").Specific
            oedit.DataBind.SetBound(True, "", "dbFind3")
            oedit.ChooseFromListUID = "CFL_5"
            oedit.ChooseFromListAlias = "OcrCode"
            oedit = objform.Items.Item("ed5").Specific
            oedit.DataBind.SetBound(True, "", "dbFind4")
            oedit.ChooseFromListUID = "CFL_6"
            oedit.ChooseFromListAlias = "OcrCode"
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTest.DoQuery("Select * from ODIM order by ""DimCode""")
            Dim ost As SAPbouiCOM.StaticText
            For intRow As Integer = 0 To oTest.RecordCount - 1
                If oTest.Fields.Item("DimActive").Value = "Y" Then
                    objform.Items.Item("ed" & intRow + 1).Visible = True
                    objform.Items.Item("st" & intRow + 1).Visible = True
                    ost = objform.Items.Item("st" & intRow + 1).Specific
                    ost.Caption = oTest.Fields.Item("DimDesc").Value
                    objform.Items.Item("ed" & intRow + 1).Enabled = True
                    Try
                        oApplication.Utilities.setEdittextvalue(objform, "ed" & intRow + 1, ststring(intRow))
                    Catch ex As Exception
                    End Try
                 
                Else
                    objform.Items.Item("ed" & intRow + 1).Visible = False
                    objform.Items.Item("st" & intRow + 1).Visible = False
                End If
                oTest.MoveNext()
            Next
            If strformid = "Approved" Then
                objform.Items.Item("3").Enabled = False
            Else
                objform.Items.Item("3").Enabled = True
            End If
            objform.Freeze(False)

        Catch ex As Exception
            oApplication.SBO_Application.MessageBox(ex.Message)
            oApplication.SBO_Application.MessageBox(ex.Message)
        Finally
        End Try
    End Sub
#End Region

#Region "Update On hand Qty"
    Private Sub FillOnhandqty(ByVal strItemcode As String, ByVal strwhs As String, ByVal aGrid As SAPbouiCOM.Grid)
        Dim oTemprec As SAPbobsCOM.Recordset
        oTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strBin, strSql As String
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strBin = aGrid.DataTable.GetValue(0, intRow)
            strSql = "Select ifnull(Sum(U_InQty)-sum(U_OutQty),0) from [@DABT_BTRN] where U_Itemcode='" & strItemcode & "' and U_FrmWhs='" & strwhs & "' and U_BinCode='" & strBin & "'"
            oTemprec.DoQuery(strSql)
            Dim dblOnhand As Double
            dblOnhand = oTemprec.Fields.Item(0).Value

            aGrid.DataTable.SetValue(2, intRow, dblOnhand.ToString)
        Next
    End Sub
#End Region

#Region "Get Form"
    '******************************************************************
    'Type               : Function
    'Name               : GetForm
    'Parameter          : FormUID
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Get The Form
    '******************************************************************
    Public Function GetForm(ByVal FormUID As String) As SAPbouiCOM.Form
        Return oApplication.SBO_Application.Forms.Item(FormUID)
    End Function
#End Region

#Region "FormDataEvent"


#End Region

#Region "Class Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

    End Sub
#End Region

#Region "getBOQReference"
    Private Function getBOQReference(ByVal aItemCode As String, ByVal aProject As String, ByVal aProcess As String, ByVal aActivity As String) As String
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select ifnull(U_Z_BOQREF,'') from [@Z_PRJ2] where U_Z_ItemCode='" & aItemCode & "' and  U_Z_PRJCODE='" & aProject.Replace("'", "''") & "' and U_Z_MODNAME='" & aProcess.Replace("'", "''") & "' and U_Z_ACTNAME='" & aActivity.Replace("'", "''") & "'")
        Return oTest.Fields.Item(0).Value
    End Function
#End Region

#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        BubbleEvent = True
        If pVal.FormTypeEx = frm_DisRule Then


            Select Case pVal.BeforeAction
                Case True
                Case False
                    Select Case pVal.EventType


                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            If pVal.ItemUID = "3" Then
                                Dim stvalue As String
                                stvalue = oApplication.Utilities.getEdittextvalue(oForm, "ed1")
                                stvalue = stvalue & ";" & oApplication.Utilities.getEdittextvalue(oForm, "ed2")
                                stvalue = stvalue & ";" & oApplication.Utilities.getEdittextvalue(oForm, "ed3")
                                stvalue = stvalue & ";" & oApplication.Utilities.getEdittextvalue(oForm, "ed4")
                                stvalue = stvalue & ";" & oApplication.Utilities.getEdittextvalue(oForm, "ed5")
                                frmSourceForm = oApplication.SBO_Application.Forms.Item(SourceFormUID)
                                Dim ststring1 As String()
                                ststring1 = stvalue.Split(";")

                                If frmSourceForm.TypeEx = frm_SalePriceBP Then
                                    Dim oGrid As SAPbouiCOM.Matrix
                                    oGrid = frmSourceForm.Items.Item(ItemUID).Specific
                                    oApplication.Utilities.SetMatrixValues(oGrid, sourceColumID, sourcerowId, stvalue)
                                    oApplication.Utilities.SetMatrixValues(oGrid, "U_Z_OcrCode", sourcerowId, ststring1(0))
                                    oApplication.Utilities.SetMatrixValues(oGrid, "U_Z_OcrCode1", sourcerowId, ststring1(1))
                                    oApplication.Utilities.SetMatrixValues(oGrid, "U_Z_OcrCode2", sourcerowId, ststring1(2))
                                    oApplication.Utilities.SetMatrixValues(oGrid, "U_Z_OcrCode3", sourcerowId, ststring1(3))
                                    oApplication.Utilities.SetMatrixValues(oGrid, "U_Z_OcrCode4", sourcerowId, ststring1(4))
                                End If
                                If frmSourceForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    frmSourceForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                                oForm.Close()
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
                                    val = oDataTable.GetValue("OcrCode", 0)
                                    oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
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
    End Sub
#End Region

End Class

