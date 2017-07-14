Public Class guia_remision

    Private XmlForm As String = Replace(Application.StartupPath & "\guia_remision.srf", "\\", "\")
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter
    Public code As String = ""
    Public oMatrix As SAPbouiCOM.Matrix
    Public oLineSelected As Integer = -1
    Public Sub New()
        MyBase.New()
        Try
            Me.SBO_Application = UDT_UF.SBOApplication
            Me.oCompany = UDT_UF.Company

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "GREMISION_") = False Then
                Dim FechaInicioTraslado As SAPbouiCOM.EditText
                Dim fechaFinTraslado As SAPbouiCOM.EditText
                Dim fechaEnvio As SAPbouiCOM.EditText
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("GREMISION_")
                oForm.Left = 400
                oForm.DataSources.UserDataSources.Add("Date", SAPbouiCOM.BoDataType.dt_DATE)
                oForm.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_DATE)
                oForm.DataSources.UserDataSources.Add("Date3", SAPbouiCOM.BoDataType.dt_DATE)
               
                oForm.Visible = True
                oMatrix = oForm.Items.Item("Item_35").Specific
                oForm.PaneLevel = 1

            Else
                oForm = Me.SBO_Application.Forms.Item("GREMISION_")
            End If
            oForm.EnableMenu("1292", True)
            oForm.EnableMenu("1293", True)
            oForm.EnableMenu("1283", False)
            oForm.EnableMenu("1284", False)
            oMatrix.ClearSelections()
            oMatrix.FlushToDataSource()
            UDT_UF.code = ""
            carcarSeries()
        Catch ex As Exception
            SBOApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Sub
    Private Sub LoadFromXML(ByRef FileName As String)
        Try
            Dim oXmlDoc As Xml.XmlDocument

            oXmlDoc = New Xml.XmlDocument

            ' ''// load the content of the XML File
            ''Dim sPath As String

            ''sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'oXmlDoc.Load(sPath & "\" & FileName)
            oXmlDoc.Load(FileName)

            '// load the form to the SBO application in one batch
            SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        If FormUID = "GREMISION_" Then
            'If pVal.ItemUID <> "Item_21" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
            'Dim txtNoGuia As SAPbouiCOM.EditText = oForm.Items.Item("Item_21").Specific
            'If txtNoGuia.Value.Trim = "" Then
            'txtNoGuia.Value = UDT_UF.code
            'End If
            ' End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                oCFLEvento = pVal
                Dim sCFL_ID As String
                sCFL_ID = oCFLEvento.ChooseFromListUID
                Dim oForm As SAPbouiCOM.Form
                oForm = SBO_Application.Forms.Item(FormUID)
                Dim oCFL As SAPbouiCOM.ChooseFromList
                oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                If oCFLEvento.BeforeAction = False Then
                    Dim oDataTable As SAPbouiCOM.DataTable
                    oDataTable = oCFLEvento.SelectedObjects
                    Dim val As String
                    If (pVal.ItemUID = "Item_9") Then
                        Try
                            Dim txtRuc As SAPbouiCOM.EditText = oForm.Items.Item("Item_9").Specific
                            ' Dim txtNomEmp As SAPbouiCOM.EditText = oForm.Items.Item("1000008").Specific
                            val = oDataTable.GetValue("CardCode", 0)
                            txtRuc.Value = val
                        Catch ex As Exception

                        End Try

                    End If
                    If (pVal.ItemUID = "Item_28") Then
                        Try
                            Dim oDoc As SAPbouiCOM.EditText = oForm.Items.Item("Item_28").Specific
                            ' Dim txtNomEmp As SAPbouiCOM.EditText = oForm.Items.Item("1000008").Specific
                            val = oDataTable.GetValue("DocEntry", 0)
                            oDoc.Value = val
                        Catch ex As Exception

                        End Try

                    End If
                    If (pVal.ItemUID = "Item_25") Then
                        Try
                            Dim oDoc As SAPbouiCOM.EditText = oForm.Items.Item("Item_25").Specific
                            Dim Name As SAPbouiCOM.EditText = oForm.Items.Item("Item_27").Specific
                            Dim Placa As SAPbouiCOM.EditText = oForm.Items.Item("Item_22").Specific
                            ' Dim txtNomEmp As SAPbouiCOM.EditText = oForm.Items.Item("1000008").Specific
                            val = oDataTable.GetValue("Code", 0)
                            Name.Value = obtenerName(val)
                            Placa.Value = obtenerPlaca(val)
                            oDoc.Value = val
                        Catch ex As Exception

                        End Try

                    End If
                End If

            End If
        End If
        If FormUID = "GREMISION_" And pVal.ItemUID = "Item_35" And pVal.Before_Action = True Then
            oLineSelected = pVal.Row
        End If
        ' If pVal.FormTypeEx = "GREMISION" And pVal.ItemUID = "Item_21" And pVal.Before_Action = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then

        ' Dim numero As New retencion_numeros
        ' BubbleEvent = False
        ' Return
        'End If
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                UDT_UF.code = ""
                oMatrix.Clear()
                oMatrix.AddRow(1)
                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            End If
            If pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                UDT_UF.code = ""
                oMatrix.AddRow(1)
                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                BubbleEvent = False
                Return
            End If
            If pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                oForm.Freeze(True)
                If oMatrix.RowCount = 0 Then
                    oMatrix.AddRow(1)
                End If
                oLineSelected = -1
                oForm.Freeze(False)
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub carcarSeries()
        Try
            Dim oRecord As SAPbobsCOM.Recordset
            Dim oSeries As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_18").Specific
            oSeries.ValidValues.LoadSeries("GREMISION", SAPbouiCOM.BoSeriesMode.sf_View)
            'oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecord.DoQuery("SP_SERIES_GUIA_REMISION")
            'If oRecord.RecordCount > 0 Then
            'While oRecord.EoF = False
            'oSeries.ValidValues.Add(oRecord.Fields.Item(0).Value, oRecord.Fields.Item(1).Value)
            'oRecord.MoveNext()
            ' End While
            'End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub addLine()
        oMatrix.AddRow(1, oMatrix.RowCount + 1)
    End Sub

    Private Function obtenerName(val As String) As String
        Dim nombre As String = ""
        Try
            Dim oRecord As SAPbobsCOM.Recordset
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("EXEC SP_OBTENER_TRANSPORTISTA '" & val & "','1'")
            nombre = oRecord.Fields.Item(0).Value
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage (ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short ,True )
        End Try
        Return nombre
    End Function

    Private Function obtenerPlaca(val As String) As String
        Dim placa As String = ""
        Try
            Dim oRecord As SAPbobsCOM.Recordset
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("EXEC SP_OBTENER_TRANSPORTISTA '" & val & "','2'")
            placa = oRecord.Fields.Item(0).Value
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()
        Catch ex As Exception

        End Try
        Return placa
    End Function


End Class
