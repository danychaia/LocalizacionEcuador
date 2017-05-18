Public Class series
    Private XmlForm As String = Replace(Application.StartupPath & "\series.srf", "\\", "\")
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter
    Public code As String = ""
    Public Sub New()
        MyBase.New()
        Try
            Me.SBO_Application = UDT_UF.SBOApplication
            Me.oCompany = UDT_UF.Company

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "frS") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("frS")
                oForm.Left = 400
                oForm.Visible = True
                oForm.PaneLevel = 1

                Dim oEstable As SAPbouiCOM.EditText
                Dim oPunto As SAPbouiCOM.EditText
                Dim oCombo As SAPbouiCOM.ComboBox
                Dim oDeI As SAPbouiCOM.EditText
                Dim oHastaI As SAPbouiCOM.EditText
                Dim oNoAutori As SAPbouiCOM.EditText
                Dim oCaducidad As SAPbouiCOM.EditText
                'Dim cmdenviar As SAPbouiCOM.Button
                'esto es para poder hacer que los textos tengan formato de fecha
                oForm.DataSources.DataTables.Add("MyDataTable")
                oForm.DataSources.UserDataSources.Add("Date", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                oForm.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                oForm.DataSources.UserDataSources.Add("De", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                oForm.DataSources.UserDataSources.Add("Hasta", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                oForm.DataSources.UserDataSources.Add("Date3", SAPbouiCOM.BoDataType.dt_DATE)
                oEstable = oForm.Items.Item("Item_2").Specific
                oPunto = oForm.Items.Item("Item_4").Specific
                oDeI = oForm.Items.Item("Item_9").Specific
                oHastaI = oForm.Items.Item("Item_11").Specific
                oNoAutori = oForm.Items.Item("Item_13").Specific
                oCaducidad = oForm.Items.Item("Item_15").Specific
                oEstable.DataBind.SetBound(True, "", "Date")
                oPunto.DataBind.SetBound(True, "", "Date2")
                oHastaI.DataBind.SetBound(True, "", "Hasta")
                oDeI.DataBind.SetBound(True, "", "De")
                oCaducidad.DataBind.SetBound(True, "", "Date3")
                oCombo = oForm.Items.Item("Item_8").Specific
                oCombo.ValidValues.Add("01", "Electrónico")
                oCombo.ValidValues.Add("02", "Impreso")               
            Else
                oForm = Me.SBO_Application.Forms.Item("frS")
            End If
            seriesImpresas()
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
        Try
            If pVal.FormTypeEx = "60004" And pVal.Before_Action = True And pVal.FormUID = "frS" Then
                If pVal.ItemUID = "Item_5" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                    Dim obutton As SAPbouiCOM.Button
                    obutton = oForm.Items.Item("Item_5").Specific
                    Dim oSerie As SAPbouiCOM.ComboBox
                    Dim oNoAutori As SAPbouiCOM.EditText
                    Dim oCaducidad As SAPbouiCOM.EditText
                    Dim oDire As SAPbouiCOM.EditText
                    Dim oCiudad As SAPbouiCOM.EditText
                    Dim oTelefono As SAPbouiCOM.EditText
                    oSerie = oForm.Items.Item("Item_24").Specific
                    oNoAutori = oForm.Items.Item("Item_13").Specific
                    oCaducidad = oForm.Items.Item("Item_15").Specific
                    oDire = oForm.Items.Item("Item_18").Specific
                    oCiudad = oForm.Items.Item("Item_21").Specific
                    oTelefono = oForm.Items.Item("Item_23").Specific
                    If obutton.Caption.Equals("Agregar") Then
                        Dim sql As String = ""
                        If oSerie.Value.Trim = "" Or oNoAutori.Value = "" Or oCaducidad.Value = "" Or oDire.Value = "" Or oCiudad.Value = "" Or oTelefono.Value = "" Then
                            SBOApplication.SetStatusBarMessage("L(821)---Debe de Ingresar toda la información ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            BubbleEvent = False
                            Return
                        End If

                        sql = "EXEC SERIES_PTO_ESTABLE '1','" & oSerie.Value.Trim & "','" & oNoAutori.Value & "','" & oCaducidad.Value & "','" & oDire.Value & "','" & oCiudad.Value & "','" & oTelefono.Value.Trim & "'"
                        Dim orecord As SAPbobsCOM.Recordset
                        orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        orecord.DoQuery(sql)
                        If orecord.Fields.Item(0).Value = "0" Then
                            SBO_Application.SetStatusBarMessage("Serie creada Correctamente", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                            limpiar()
                        Else
                            SBO_Application.SetStatusBarMessage(orecord.Fields.Item(0).Value.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End If

                        carcarSeries()
                        seriesImpresas()
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                        orecord = Nothing
                        GC.Collect()
                    Else
                        If obutton.Caption.Equals("Eliminar") Then
                            If code <> "" Then
                                Dim orecord As SAPbobsCOM.Recordset
                                Dim Sql As String
                                orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Sql = "EXEC SERIES_PTO_ESTABLE '3','" & code & "','" & oNoAutori.Value & "','" & oCaducidad.Value & "','" & oDire.Value & "','" & oCiudad.Value & "','" & oTelefono.Value.Trim & "'"
                                orecord.DoQuery(sql)
                                carcarSeries()
                                seriesImpresas()
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                                orecord = Nothing
                                GC.Collect()
                            Else
                                SBO_Application.SetStatusBarMessage("Debe de seleccionar una fila", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            End If
                        End If
                    End If
                
            End If
            If pVal.ItemUID = "Item_0" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim gridView As SAPbouiCOM.Grid
                gridView = oForm.Items.Item("Item_0").Specific
                If pVal.Row <> -1 Then
                    code = gridView.DataTable.GetValue("Code", pVal.Row).ToString
                    Dim obutton As SAPbouiCOM.Button
                    obutton = oForm.Items.Item("Item_5").Specific
                    obutton.Caption = "Eliminar"
                    UDT_UF.docEntry = docEntry
                    'Dim detalle As New retencion_info_detalle
                    BubbleEvent = False
                End If

            End If
                If (pVal.ItemUID = "Item_24" Or pVal.ItemUID = "Item_4") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Before_Action = True Then
                    Dim obutton As SAPbouiCOM.Button
                    obutton = oForm.Items.Item("Item_5").Specific
                    obutton.Caption = "Agregar"
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message)
        End Try
       
    End Sub
    Private Sub carcarSeries()
        Try
            Dim gridView As SAPbouiCOM.Grid
            gridView = oForm.Items.Item("Item_0").Specific
            gridView.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            Dim sql As String = "EXEC [SERIES_PTO_ESTABLE] '2','','','','','',''"
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(sql)
            gridView.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            gridView.AutoResizeColumns()
            gridView.Columns.Item(0).Visible = False
            gridView.Columns.Item(1).Visible = False
            gridView.Columns.Item(2).Editable = False
            gridView.Columns.Item(3).Editable = False
            gridView.Columns.Item(4).Editable = False
            gridView.Columns.Item(5).Editable = False
            gridView.Columns.Item(6).Editable = False
            gridView.Columns.Item(7).Editable = False
           
            System.Runtime.InteropServices.Marshal.ReleaseComObject(gridView)
            gridView = Nothing
            GC.Collect()
        Catch ex As Exception
            Me.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
        End Try
    End Sub

    Private Sub limpiar()
        Try
            Dim oDe As SAPbouiCOM.EditText
            Dim oHasta As SAPbouiCOM.EditText
            Dim oDeI As SAPbouiCOM.EditText
            Dim oHastaI As SAPbouiCOM.EditText
            Dim oCombo As SAPbouiCOM.ComboBox
            Dim oNoAutori As SAPbouiCOM.EditText
            Dim oCaducidad As SAPbouiCOM.EditText
            Dim oDocumen As SAPbouiCOM.ComboBox
            Dim oDire As SAPbouiCOM.EditText
            Dim oCiudad As SAPbouiCOM.EditText
            Dim oTelefono As SAPbouiCOM.EditText
            oDe = oForm.Items.Item("Item_2").Specific
            oHasta = oForm.Items.Item("Item_4").Specific
            oCombo = oForm.Items.Item("Item_8").Specific
            oDeI = oForm.Items.Item("Item_9").Specific
            oHastaI = oForm.Items.Item("Item_11").Specific
            oNoAutori = oForm.Items.Item("Item_13").Specific
            oCaducidad = oForm.Items.Item("Item_15").Specific
            oDocumen = oForm.Items.Item("Item_19").Specific
            oDire = oForm.Items.Item("Item_18").Specific
            oCiudad = oForm.Items.Item("Item_21").Specific
            oTelefono = oForm.Items.Item("Item_23").Specific

            oDe.Value = ""
            oHasta.Value = ""
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oDeI.Value = ""
            oHastaI.Value = ""
            oNoAutori.Value = ""
            oCaducidad.Value = ""
            oDocumen.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oDire.Value = ""
            oCiudad.Value = ""
            oTelefono.Value = ""


        Catch ex As Exception

        End Try

    End Sub

    Private Sub seriesImpresas()
        Try
            Dim oRecord As SAPbobsCOM.Recordset
            Dim oComboSeries As SAPbouiCOM.ComboBox
            oComboSeries = oForm.Items.Item("Item_24").Specific
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("EXEC SP_CARGAR_SERIES ")
            If oRecord.RecordCount > 0 Then
                For i As Integer = oComboSeries.ValidValues.Count - 1 To 0 Step -1
                    oComboSeries.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
                While oRecord.EoF = False
                    oComboSeries.ValidValues.Add(oRecord.Fields.Item(0).Value, oRecord.Fields.Item(1).Value)
                    oRecord.MoveNext()
                End While
            End If
        Catch ex As Exception

        End Try
    End Sub

End Class
