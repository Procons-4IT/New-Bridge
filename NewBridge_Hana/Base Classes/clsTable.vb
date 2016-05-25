Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try
            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab
                oUserTablesMD.TableDescription = strDesc
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            oUserTablesMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "OADM" Or strTab = "OPCH" Or strTab = "OITM" Or strTab = "OJDT" Or strTab = "INV1" Or strTab = "RDR1" Or strTab = "OINV" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "OWHS" Or strTab = "OHEM") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                oUserFieldMD.Description = strDesc
                oUserFieldMD.Name = strCol
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory

                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    '  MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
            Else


            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()

        End Try


    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE Upper(""TableID"") = '" & Table.ToString.ToUpper.Trim() & "' AND Upper(""AliasID"") = '" & Column.ToString.ToUpper.Trim() & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)
            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

            If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                '// Set the table name and the key name
                oUserKeysMD.TableName = strTab
                oUserKeysMD.KeyName = strKey

                '// Set the column's alias
                oUserKeysMD.Elements.ColumnAlias = strColumn
                oUserKeysMD.Elements.Add()
                oUserKeysMD.Elements.ColumnAlias = "RentFac"

                '// Determine whether the key is unique or not
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                '// Add the key
                If oUserKeysMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
            oUserKeysMD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal strChildTb2 As String = "", _
                                        Optional ByVal strChildTb3 As String = "", Optional ByVal strChildTb4 As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document, Optional ByVal blnCanArchive As Boolean = False, Optional ByVal strLogName As String = "")

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.LogTableName = ""
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                Dim intTables As Integer = 0
                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If
                If strChildTb2 <> "" Then
                    If strChildTbl <> "" Then
                        oUserObjectMD.ChildTables.Add()
                        intTables = intTables + 1
                    End If
                    oUserObjectMD.ChildTables.SetCurrentLine(intTables)
                    oUserObjectMD.ChildTables.TableName = strChildTb2
                End If
                If strChildTb3 <> "" Then
                    If strChildTb2 <> "" Then
                        oUserObjectMD.ChildTables.Add()
                        intTables = intTables + 1
                    End If
                    oUserObjectMD.ChildTables.SetCurrentLine(intTables)

                    oUserObjectMD.ChildTables.TableName = strChildTb3
                End If
                If strChildTb4 <> "" Then
                    If strChildTb3 <> "" Then
                        oUserObjectMD.ChildTables.Add()
                        intTables = intTables + 1
                    End If
                    oUserObjectMD.ChildTables.SetCurrentLine(intTables)
                    oUserObjectMD.ChildTables.TableName = strChildTb4
                End If

                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable


                If blnCanArchive Then
                    oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.LogTableName = strLogName
                End If

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub

    Public Function UDOExpances(ByVal strUDO As String, _
                       ByVal strDesc As String, _
                           ByVal strTable As String, _
                               ByVal intFind As Integer, _
                                   Optional ByVal strCode As String = "", _
                                       Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_ExpName"
                oUserObjects.FormColumns.FormColumnDescription = "U_ExpName"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Active"
                oUserObjects.FormColumns.FormColumnDescription = "U_Active"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_PostType"
                oUserObjects.FormColumns.FormColumnDescription = "U_PostType"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_CrActCode"
                oUserObjects.FormColumns.FormColumnDescription = "U_CrActCode"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_DbActCode"
                oUserObjects.FormColumns.FormColumnDescription = "U_DbActCode"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function
#End Region

#Region "Public Functions"
    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************
    Public Sub CreateTables()
        Dim oProgressBar As SAPbouiCOM.ProgressBar
        Try

            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            AddFields("OPCH", "Z_ARInvoice", "A/R Invoice Ref.No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OPCH", "Z_BaseEntry", "A/R Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OINV", "Z_APInvoice", "A/P Invoice Ref.No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OCRD", "Z_ComRePay", "Commission/Rebate Payable", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            AddFields("SPP1", "Z_DisRule", "Dist.Rule", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("SPP1", "Z_Comm", "Commission", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("SPP1", "Z_MarkReb", "Marketing Rebate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("SPP1", "Z_OcrCode", "Distribution Rule", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("SPP1", "Z_OcrCode1", "Costing Code 2", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("SPP1", "Z_OcrCode2", "Costing Code 3", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("SPP1", "Z_OcrCode3", "Costing Code 4", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("SPP1", "Z_OcrCode4", "Costing Code 5", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            AddTables("Z_OCRE", "Customer Rebate", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OCRE", "Z_MarkReb", "Marketing Rebate  Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OCRE", "Z_ProReb", "Commission  Rebate Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            addField("INV1", "Z_IsComm", "Commission Applied", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("INV1", "Z_Comm", "Commission", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Rate)
            AddFields("INV1", "Z_MarkReb", "Marketing Rebate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Rate)

            AddFields("INV1", "Z_Comm_Per", "Commission Percentage", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("INV1", "Z_MarkReb_Per", "Marketing Rebate Percentage", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)

            addField("OWHS", "Z_Type", "Warehouse Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "H,D,R", "HUB,DropShip,Regular", "R")
            AddFields("OITM", "Z_RebValue", "Rebate Percentage", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)

            'For Supplier Dropship Rebate Calcualtion
            AddFields("INV1", "Z_Rebate", "Rebate %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            'AddFields("OWHS", "Z_PWhs", "Physical Warehouse ", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            'HUB Rebate

            addField("SPP1", "Z_RegStatus", "Registration Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "R,N,L,O", "Registered,NRR,Pre-Licensed,None", "O")
            AddFields("OITM", "Z_SupCom", "Suppler Commission", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("OITM", "Z_SupComPre", "Suppler Commission Pre-License", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            addField("INV1", "Z_RegStatus", "Registration Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "R,N,L,O", "Registered,NRR,Pre-Licensed,None", "O")
            AddFields("RDR1", "Z_SupCom1", "Suppler Commission %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Rate)

            AddFields("INV1", "Z_ItemCost", "Item Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("INV1", "Z_Accrual", "Rebate Accrual", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)


            AddTables("Z_OSRE", "Suppler Rebate", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OSRE", "Z_COGS", "COGS  Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OSRE", "Z_Accrual", "Accrual Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OSRE", "Z_ItmsGrp", "Item Group", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)

            AddFields("OJDT", "Z_ARInvoice", "A/R Invoice Ref.No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OJDT", "Z_BaseEntry", "A/R Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OPCH", "Z_JournalRef", "Journal Entry Ref.No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OCRD", "Z_ComBase", "Commission Base %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)


            AddTables("Z_NBLOGIN", "New Bridge Login Details", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_NBLOGIN", "UID", "User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_NBLOGIN", "PWD", "Password", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_NBLOGIN", "EMPID", "Employee ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_NBLOGIN", "EMPNAME", "Employee NAME", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_NBLOGIN", "ESSAPPROVER", "ESS Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "E,M", "Employee,Manager", "E")

            AddTables("Z_NBEXPANCES", "New Bridge Expences Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddFields("Z_NBEXPANCES", "ExpName", "Expences Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_NBEXPANCES", "CrActCode", "Credit Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXPANCES", "DbActCode", "Debit Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            addField("@Z_NBEXPANCES", "Active", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            addField("@Z_NBEXPANCES", "PostType", "Posting Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "B,G", "Business Partners,G/L Account", "B")
            AddFields("Z_NBEXPANCES", "GLDesc", "GL Credit Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_NBEXPANCES", "GLDesc1", "GL Debit Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_NBEXPANCES", "Category", "Category", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "B,P", "Business Travel,Purchase Requisition", "B")


            AddTables("Z_NBOAPPT", "Approval Template", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("Z_NBAPPT1", "Approval Orginator", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("Z_NBAPPT2", "Approval Authorizer", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("Z_NBOAPPT", "Code", "Approval Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBOAPPT", "Name", "Approval Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_NBOAPPT", "DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBOAPPT", "DocDesc", "Document Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_NBOAPPT", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            addField("@Z_NBOAPPT", "Condition", "Condition", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "1,2,3", "<,>=,>", "1")
            AddFields("Z_NBOAPPT", "Amount", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)

            AddFields("Z_NBAPPT1", "OUser", "Orginator Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBAPPT1", "OName", "Orginator Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddFields("Z_NBAPPT2", "AUser", "Authorizer Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBAPPT2", "AName", "Authorizer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_NBAPPT2", "AMan", "Mandatory", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_NBAPPT2", "AFinal", "Final Stage", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            AddTables("Z_NBAPHIS", "Approval History", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_NBAPHIS", "DocEntry", "Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBAPHIS", "LineId", "Document LineId", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBAPHIS", "DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBAPHIS", "EmpId", "Employee Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBAPHIS", "EmpName", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_NBAPHIS", "AppStatus", "Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,A,R", "Pending,Approved,Rejected", "P")
            AddFields("Z_NBAPHIS", "Remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_NBAPHIS", "ApproveBy", "Approved By", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBAPHIS", "Approvedt", "Approver Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_NBAPHIS", "ADocEntry", "Template DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_NBAPHIS", "ALineId", "Template LineId", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_NBAPHIS", "Month", "Month", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_NBAPHIS", "Year", "Year", SAPbobsCOM.BoFieldTypes.db_Numeric)

            AddFields("OHEM", "Z_NBCardCode", "NewBridge Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            ''Expenses
            AddTables("Z_NBOEXP", "Expenses Entry Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_NBOEXP", "EmpCode", "Employee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBOEXP", "EmpName", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 150)
            addField("@Z_NBOEXP", "DocStatus", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "O,C", "Opened,Closed", "O")
            AddFields("Z_NBOEXP", "LFANo", "LFA Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBOEXP", "PONo", "PO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBOEXP", "Dim1", "Dimension 1", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBOEXP", "Dim2", "Dimension 2", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBOEXP", "Dim3", "Dimension 3", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBOEXP", "Country", "Country", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBOEXP", "Product", "Product", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBOEXP", "DeptName", "Department Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            addField("@Z_NBOEXP", "TypeofExp", "Expense Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,L", "Personal,LFA", "L")
            addField("@Z_NBOEXP", "ExpType", "Expense Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,L", "Personal,LFA", "L")
            AddFields("Z_NBOEXP", "AppStatus", "Approver Status", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)

            AddTables("Z_NBEXP1", "Expenses Entry Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_NBEXP1", "ReqDate", "Request Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_NBEXP1", "ExpCode", "Expenses Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_NBEXP1", "ExpName", "Expenses Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 150)
            AddFields("Z_NBEXP1", "TransCur", "Transaction Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_NBEXP1", "TransAmt", "Transaction Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_NBEXP1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_NBEXP1", "Ref1", "Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 150)
            AddFields("Z_NBEXP1", "Attachment", "Attachment", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_NBEXP1", "ApproveId", "Approve Template Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "ApproveBy", "Approved By", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_NBEXP1", "Approvedt", "Approver Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("@Z_NBEXP1", "AppStatus", "Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,R,A", "Pending,Rejected,Approved", "P")
            AddFields("Z_NBEXP1", "CurApprover", "Current Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "NxtApprover", "Next Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            addField("@Z_NBEXP1", "AppRequired", "Approval Required", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            AddFields("Z_NBEXP1", "AppReqDate", "Required Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_NBEXP1", "ReqTime", "Required Time", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "RejRemark", "Rejection Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_NBEXP1", "CrActCode", "Credit Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "DbActCode", "Debit Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            addField("@Z_NBEXP1", "PostType", "Posting Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "B,G", "Business Partners,G/L Account", "B")
            addField("@Z_NBEXP1", "IsDeleted", "Temporary deleted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_NBEXP1", "CardCode", "NewBridge Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_NBEXP1", "DocRefCode", "Document Ref.Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "LFANo", "LFA Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "PONo", "PO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "Dim1", "Dimension 1", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "Dim2", "Dimension 2", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "Dim3", "Dimension 3", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "Country", "Country", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "Product", "Product", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "BPCurrency", "BP Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "SAPCurrency", "SAP Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "LocCurrency", "Local Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_NBEXP1", "ExcRate", "Exchange Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_NBEXP1", "LocAmount", "Local Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)

            AddTables("Z_OBUDDF", "Budget Definition", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OBUDDF", "Year", "Year", SAPbobsCOM.BoFieldTypes.db_Alpha, , 4)
            AddFields("Z_OBUDDF", "Category", "Category", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OBUDDF", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)
            addField("Z_OBUDDF", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "Y")

            AddTables("Z_BUDDF1", "Budget Definition-Lines1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_BUDDF1", "Category", "Category", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_BUDDF1", "DisCode", "Dimension", SAPbobsCOM.BoFieldTypes.db_Alpha, , 204)
            AddFields("Z_BUDDF1", "OcrCode", "Dimension 1", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_BUDDF1", "OcrCode2", "Dimension 2", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_BUDDF1", "OcrCode3", "Dimension 3", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_BUDDF1", "OcrCode4", "Dimension 4", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_BUDDF1", "OcrCode5", "Dimension 5", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_BUDDF1", "Z_Budget", "Allocated Budget", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_BUDDF1", "Z_PRApprvd", "PR Approved", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_BUDDF1", "Z_POApprvd", "PO Approved", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_BUDDF1", "Z_GRApprvd", "Good Receipt Approved", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_BUDDF1", "Z_IVApprvd", "Invoiced Approved", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_BUDDF1", "Z_ABudget", "Available Budget", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)

            CreateUDO()
        Catch ex As Exception
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Sub CreateUDO()
        Try
            AddUDO("Z_NBOEXP", "Expenses Entry", "Z_NBOEXP", "DocEntry", "U_EmpCode", "Z_NBEXP1", , , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_NBLOGIN", "LoginSetup", "Z_NBLOGIN", "DocEntry", "U_UID", , , , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_NBAPHIS", "Approval History", "Z_NBAPHIS", "DocEntry", "U_DocEntry", , , , , SAPbobsCOM.BoUDOObjType.boud_Document, True, "AZ_NBAPHIS")
            AddUDO("Z_NBOAPPT", "Approval Template", "Z_NBOAPPT", "DocEntry", "U_Code", "Z_NBAPPT1", "Z_NBAPPT2", , , SAPbobsCOM.BoUDOObjType.boud_Document)
            UDOExpances("Z_NBEXPANCES", "Expences - Master", "Z_NBEXPANCES", 1, "U_ExpName")
            AddUDO("Z_OBUDDF", "Budget Definition", "Z_OBUDDF", "DocEntry", "U_Year", "Z_BUDDF1", "", , , SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
