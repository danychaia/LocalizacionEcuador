Module UDT_UF
    Private _SBOApplication As SAPbouiCOM.Application
    Public Property SBOApplication() As SAPbouiCOM.Application
        Get
            Return _SBOApplication
        End Get
        Set(ByVal value As SAPbouiCOM.Application)
            _SBOApplication = value
        End Set
    End Property

    Private _Company As SAPbobsCOM.Company
    Public docEntry As String
    Public code As String = ""

    Public Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
        Set(ByVal value As SAPbobsCOM.Company)
            _Company = value
        End Set
    End Property

    Public Sub userField(ByVal oCompany As SAPbobsCOM.Company, ByVal tableName As String, ByVal Descripcion As String, ByVal size As Integer, ByVal namefield As String, ByVal type As SAPbobsCOM.BoFieldTypes, ByVal validation As Boolean, ByVal SBO_app As SAPbouiCOM.Application)
        Dim err As String = ""
        Dim num As Integer = 0
        Dim row As Integer = -1
        Try
            If fieldExist(oCompany, tableName, namefield) = False Then
                Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
                oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                oUserFieldsMD.TableName = tableName
                oUserFieldsMD.Name = namefield   '"DOCUMENTO"
                oUserFieldsMD.Description = Descripcion  '"DOCUMENTO"
                oUserFieldsMD.Type = type
                If type = 2 Or type = 0 Then
                    oUserFieldsMD.EditSize = size
                End If

                If validation = True Then
                    oUserFieldsMD.ValidValues.Value = "1"
                    oUserFieldsMD.ValidValues.Description = "INICIO"
                    oUserFieldsMD.ValidValues.Add()
                End If
                If oUserFieldsMD.Add() <> 0 Then
                    oCompany.GetLastError(num, err)
                    SBO_app.SetStatusBarMessage(num & " " & err, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            End If

            GC.Collect()
        Catch ex As Exception
            SBO_app.SetStatusBarMessage(ex.Message & "  " & num & " " & err, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try


    End Sub


    Public Sub userTable(ByVal oCompany As SAPbobsCOM.Company, ByVal tableName As String, ByVal Descripcion As String, ByVal size As Integer, ByVal namefield As String, ByVal type As SAPbobsCOM.BoUTBTableType, ByVal validation As Boolean, ByVal SBO_app As SAPbouiCOM.Application)
        Try

            Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
            Dim iResult As Long
            Dim sMsg As String
            Dim sTable As String

            Try
                oUserTablesMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                Dim tabla = "@" & tableName
                If (oUserTablesMD.GetByKey(tabla) = False) Then

                    oUserTablesMD.TableName = tableName
                    oUserTablesMD.TableDescription = Descripcion
                    oUserTablesMD.TableType = type
                    oUserTablesMD.Add()

                    oUserTablesMD.Update()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
                    oUserTablesMD = Nothing
                    GC.Collect()
                End If


            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        Catch ex As Exception
            SBO_app.SetStatusBarMessage(ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Long, True)
        End Try
    End Sub

    Private Function fieldExist(oCompany As SAPbobsCOM.Company, tableName As String, namefield As String) As Boolean

        Dim existe As Boolean = False
        Dim record As SAPbobsCOM.Recordset

        record = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        record.DoQuery("SELECT a.AliasID   FROM CUFD a WHERE TableID = '" & tableName & "' AND AliasID = '" & namefield & "'")
        If record.RecordCount > 0 Then
            existe = True
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(record)
        record = Nothing
        GC.Collect()
        Return existe
    End Function

    Public Function ActivateFormIsOpen(ByVal SboApplication As SAPbouiCOM.Application, ByVal FormID As String) As Boolean
        Try
            Dim result As Boolean = False
            For x = 0 To SboApplication.Forms.Count - 1
                If SboApplication.Forms.Item(x).UniqueID = FormID Then
                    SboApplication.Forms.Item(x).Select()
                    result = True
                    Exit For
                End If
            Next
            Return result
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function


    Public Function updateUserField(oCompany As SAPbobsCOM.Company, tableName As String, namefield As String, validArray As ArrayList) As Boolean

        Dim existe As Boolean = False
        Dim record As SAPbobsCOM.Recordset

        Dim eler As Integer = 0
        Dim mensa As String = ""

        record = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql = "SELECT a.FieldID   FROM CUFD a WHERE TableID = '" & tableName & "' AND AliasID = '" & namefield & "'"
        record.DoQuery(sql)
        If record.RecordCount > 0 Then

            Dim oFielID = record.Fields.Item("FieldID").Value
            System.Runtime.InteropServices.Marshal.ReleaseComObject(record)
            record = Nothing
            GC.Collect()
            Dim oUserField As SAPbobsCOM.UserFieldsMD
            oUserField = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            If oUserField.GetByKey(tableName, oFielID) Then

                For Each lista As validValues In validArray
                    oUserField.ValidValues.Value = lista.value
                    oUserField.ValidValues.Description = lista.descrip
                    oUserField.ValidValues.Add()
                Next
                If oUserField.Update() <> 0 Then
                    oCompany.GetLastError(eler, mensa)
                    SBOApplication.SetStatusBarMessage(mensa, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If
            End If

        End If



        Return existe
    End Function

  

End Module
