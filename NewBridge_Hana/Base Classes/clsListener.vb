Public Class clsListener
    Inherits Object
    Private ThreadClose As New Threading.Thread(AddressOf CloseApp)
    Private WithEvents _SBO_Application As SAPbouiCOM.Application
    Private _Company As SAPbobsCOM.Company
    Private _Utilities As clsUtilities
    Private _Collection As Hashtable
    Private _LookUpCollection As Hashtable
    Private _FormUID As String
    Private _Log As clsLog_Error
    Private oMenuObject As Object
    Private oItemObject As Object
    Private oSystemForms As Object
    Dim objFilters As SAPbouiCOM.EventFilters
    Dim objFilter As SAPbouiCOM.EventFilter

#Region "New"
    Public Sub New()
        MyBase.New()
        Try
            _Company = New SAPbobsCOM.Company
            _Utilities = New clsUtilities
            _Collection = New Hashtable(10, 0.5)
            _LookUpCollection = New Hashtable(10, 0.5)
            oSystemForms = New clsSystemForms
            _Log = New clsLog_Error

            SetApplication()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Public Properties"

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return _SBO_Application
        End Get
    End Property

    Public ReadOnly Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
    End Property

    Public ReadOnly Property Utilities() As clsUtilities
        Get
            Return _Utilities
        End Get
    End Property

    Public ReadOnly Property Collection() As Hashtable
        Get
            Return _Collection
        End Get
    End Property

    Public ReadOnly Property LookUpCollection() As Hashtable
        Get
            Return _LookUpCollection
        End Get
    End Property

    Public ReadOnly Property Log() As clsLog_Error
        Get
            Return _Log
        End Get
    End Property
#Region "Filter"

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        oApplication.SetFilter(Filters)
    End Sub
    Public Sub SetFilter()
        Try
            ''Form Load
            objFilters = New SAPbouiCOM.EventFilters

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            objFilter.AddEx(frm_SalesOrder)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            objFilter.Add(frm_SalesOrder)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region

#End Region

#Region "Menu Event"

    Private Sub _SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.FormDataEvent
        Select Case BusinessObjectInfo.FormTypeEx
            Case frm_Invoice, frm_ARCreditNote
                Dim objInvoice As clsRebatePosting
                objInvoice = New clsRebatePosting
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            Case frm_APInvoice
                Dim objInvoice As clsSupplierRebate
                objInvoice = New clsSupplierRebate
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_S01
                Dim objInvoice As clsNBApprovalTemplate
                objInvoice = New clsNBApprovalTemplate
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
        End Select
        '  End If
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_Posting
                        oMenuObject = New clsMissingRebatePosting
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_AppTemp
                        oMenuObject = New clsNBApprovalTemplate
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Logsetup
                        oMenuObject = New clsNBLogin
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Expenses
                        oMenuObject = New clsNBExpenses
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_CustRebate
                        oMenuObject = New clsCustCommission
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_SubRebate
                        oMenuObject = New clsSupplierCommission
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_OBUDDF
                        oMenuObject = New clsBudgetDefinition
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Z_BUD_R
                        oMenuObject = New clsBudgetReport
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ADD, mnu_FIND, mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD_ROW, mnu_DELETE_ROW
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If

                End Select

            Else
                Select Case pVal.MenuUID
                    Case mnu_Posting
                        oMenuObject = New clsMissingRebatePosting
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_CLOSE
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If
                    Case mnu_ADD, mnu_FIND, mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD_ROW, mnu_DELETE_ROW
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If
                End Select

            End If

        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oMenuObject = Nothing
        End Try
    End Sub
#End Region

#Region "Item Event"
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ItemEvent
        Try
            _FormUID = FormUID

            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                Select Case pVal.FormType
                End Select
            End If

            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                Select Case pVal.FormTypeEx
                    Case frm_Posting
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsMissingRebatePosting
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_S01
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsNBApprovalTemplate
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Expenses
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsNBExpenses
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_LoginSetup
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsNBLogin
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_CustComDef
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCustCommission
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_SubComDef
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSupplierCommission
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_SalePriceBP
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSalesPriceBP
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_DisRule
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDisRule
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Invoice, frm_ARCreditNote
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsRebatePosting
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_APInvoice
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSupplierRebate
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_OBUDDF
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsBudgetDefinition
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Z_BUD_R
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsBudgetReport
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                End Select
            End If

            If _Collection.ContainsKey(FormUID) Then
                oItemObject = _Collection.Item(FormUID)
                If oItemObject.IsLookUpOpen And pVal.BeforeAction = True Then
                    _SBO_Application.Forms.Item(oItemObject.LookUpFormUID).Select()
                    BubbleEvent = False
                    Exit Sub
                End If
                _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.BeforeAction = False Then
                If _LookUpCollection.ContainsKey(FormUID) Then
                    oItemObject = _Collection.Item(_LookUpCollection.Item(FormUID))
                    If Not oItemObject Is Nothing Then
                        oItemObject.IsLookUpOpen = False
                    End If
                    _LookUpCollection.Remove(FormUID)
                End If

                If _Collection.ContainsKey(FormUID) Then
                    _Collection.Item(FormUID) = Nothing
                    _Collection.Remove(FormUID)
                End If

            End If

        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Application Event"
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles _SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    _Utilities.AddRemoveMenus("RemoveMenus.xml")
                    CloseApp()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Try
    End Sub
#End Region

#Region "Close Application"
    Private Sub CloseApp()
        Try
            If Not _SBO_Application Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_SBO_Application)
            End If

            If Not _Company Is Nothing Then
                If _Company.Connected Then
                    _Company.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company)
            End If

            _Utilities = Nothing
            _Collection = Nothing
            _LookUpCollection = Nothing

            ThreadClose.Sleep(10)
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Throw ex
        Finally
            oApplication = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Set Application"
    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                SboGuiApi = New SAPbouiCOM.SboGuiApi
                SboGuiApi.Connect(sConnectionString)
                _SBO_Application = SboGuiApi.GetApplication()
            Else
                Throw New Exception("Connection string missing.")
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SboGuiApi = Nothing
        End Try
    End Sub
#End Region

#Region "Finalize"
    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            '            CloseApp()

            oMenuObject = Nothing
            oItemObject = Nothing
            oSystemForms = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Addon Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

End Class
