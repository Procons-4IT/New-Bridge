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

            If Not (strTab = "OADM" Or strTab = "OPCH" Or strTab = "OITM" Or strTab = "OJDT" Or strTab = "INV1" Or strTab = "RDR1" Or strTab = "OINV" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "OWHS") Then
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
                                        Optional ByVal strChildTbl As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document)

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
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

                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If

                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

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
        Catch ex As Exception
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Sub CreateUDO()
        Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class
