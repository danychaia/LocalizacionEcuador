Imports System.Xml

Public Class fact_compro
    Private XmlForm As String = Replace(Application.StartupPath & "\factura_comprobante.srf", "\\", "\")
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter

    Public Sub New()
        MyBase.New()
        Try
            Me.SBO_Application = UDT_UF.SBOApplication
            Me.oCompany = UDT_UF.Company

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "frmCfac") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("frmCfac")
                oForm.Visible = True
                oForm.PaneLevel = 1
                oForm.Left = 419
                Dim inicio As SAPbouiCOM.EditText
                Dim fin As SAPbouiCOM.EditText
                Dim cmdenviar As SAPbouiCOM.Button
                'esto es para poder hacer que los textos tengan formato de fecha
                oForm.DataSources.DataTables.Add("MyDataTable")
                oForm.DataSources.UserDataSources.Add("Date", SAPbouiCOM.BoDataType.dt_DATE)
                oForm.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_DATE)
                inicio = oForm.Items.Item("txtInicio").Specific
                fin = oForm.Items.Item("txtfin").Specific
                inicio.DataBind.SetBound(True, "", "Date")
                fin.DataBind.SetBound(True, "", "Date2")
                cargarSeries()
            Else
                oForm = Me.SBO_Application.Forms.Item("frmCfac")
            End If

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

    Private Sub cargarSeries()
        Try
            Dim seriesCombo As SAPbouiCOM.ComboBox
            Dim orecord As SAPbobsCOM.Recordset
            orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orecord.DoQuery("exec series")
            seriesCombo = oForm.Items.Item("Item_1").Specific
            If orecord.RecordCount > 0 Then
                While orecord.EoF = False
                    seriesCombo.ValidValues.Add(orecord.Fields.Item(0).Value, orecord.Fields.Item(1).Value.ToString & "(" & orecord.Fields.Item(2).Value.ToString & ")")
                    orecord.MoveNext()
                End While
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
            orecord = Nothing
            GC.Collect()
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormTypeEx = "60004" And pVal.Before_Action = True And pVal.FormUID = "frmCfac" Then
                If pVal.ItemUID = "btn_Buscar" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                    Dim ini As SAPbouiCOM.EditText
                    Dim fin As SAPbouiCOM.EditText
                    Dim com As SAPbouiCOM.ComboBox
                    com = oForm.Items.Item("Item_1").Specific
                    ini = oForm.Items.Item("txtInicio").Specific
                    fin = oForm.Items.Item("txtfin").Specific
                    If com.Value.ToString = "" Then
                        Me.SBO_Application.SetStatusBarMessage("Debe Seleccionar una serie")
                        BubbleEvent = False
                    Else
                        If ini.Value.ToString = "" Then
                            Me.SBO_Application.SetStatusBarMessage("Debe Seleccionar una fecha de inicio")
                            BubbleEvent = False
                        Else
                            If fin.Value = "" Then
                                Me.SBO_Application.SetStatusBarMessage("Debe de seleccionar una fecha de Final")
                                BubbleEvent = False
                            Else
                                buscar(ini.Value, fin.Value)
                            End If
                        End If
                    End If

                    BubbleEvent = False
                End If


                If pVal.ItemUID = "bt_pro" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    generarXML()
                End If
            End If

        Catch ex As Exception
            Me.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
    End Sub

    Private Sub buscar(p1 As String, p2 As String)
        Try
            Dim oRecord As SAPbobsCOM.Recordset
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sql As String = "BUSCAR_FACTURA_FACE '" & p1 & "','" & p2 & "'"
            oRecord.DoQuery(sql)
            If oRecord.RecordCount > 0 Then
                Dim gridView As SAPbouiCOM.Grid
                gridView = oForm.Items.Item("txtGridFac").Specific
                oForm.DataSources.DataTables.Item(0).ExecuteQuery(sql)
                gridView.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
                gridView.AutoResizeColumns()
                gridView.Columns.Item(0).Visible = False
                gridView.Columns.Item(1).Editable = False
                gridView.Columns.Item(2).Editable = False
                gridView.Columns.Item(3).Editable = False
                gridView.Columns.Item(4).Editable = False
                gridView.Columns.Item(5).Editable = False
                gridView.Columns.Item(6).Editable = False
                gridView.Columns.Item(7).Editable = False
                gridView.Columns.Item(8).Editable = False
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                oRecord = Nothing
                GC.Collect()
                Dim boton As SAPbouiCOM.Button
                boton = oForm.Items.Item("bt_pro").Specific
                If gridView.Rows.Count > 0 Then
                    boton.Item.Enabled = True
                Else
                    boton.Item.Enabled = False
                End If
            Else
                Me.SBO_Application.SetStatusBarMessage("No se encontraron datos", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If

        Catch ex As Exception
            Me.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub generarXML()
        Try
            Dim grd As SAPbouiCOM.Grid
            Dim doc As New XmlDocument
            Dim Nodo As XmlNode
            Dim oRecord As SAPbobsCOM.Recordset
            Dim oProgressive As SAPbouiCOM.ProgressBar
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            grd = oForm.Items.Item("txtGridFac").Specific
            If grd.Rows.Count > 0 Then
                oProgressive = SBO_Application.StatusBar.CreateProgressBar("Generando Retencion de :", grd.Rows.Count, True)
                For i = 0 To grd.Rows.Count - 1
                    Dim docEntry = grd.DataTable.GetValue(grd.DataTable.Columns.Item(0).Name, i)
                    oRecord.DoQuery("exec ENCABEZADO_FACTURA '" & docEntry & "'")
                    Dim writer As New XmlTextWriter("Comprobante (F) No." & docEntry.ToString & ".xml", System.Text.Encoding.UTF8)
                    writer.WriteStartDocument(True)
                    writer.Formatting = Formatting.Indented
                    writer.Indentation = 2
                    writer.WriteStartElement("factura")
                    writer.WriteAttributeString("id", "comprobante")
                    writer.WriteAttributeString("version", "2.0.0")
                    writer.WriteStartElement("infoTributaria")
                    createNode("ambiente", oRecord.Fields.Item(0).Value.ToString, writer)
                    createNode("tipoEmision", oRecord.Fields.Item(1).Value.ToString, writer)
                    createNode("razonSocial", oRecord.Fields.Item(2).Value.ToString, writer)
                    createNode("ruc", oRecord.Fields.Item(3).Value.ToString.PadLeft(13, "0"), writer)
                    createNode("claveAcesso", claveAcceso(oRecord).PadLeft(49, "0"), writer)
                    createNode("codDoc", oRecord.Fields.Item("codDoc").Value.ToString.PadLeft(2, "0"), writer)
                    createNode("estab", oRecord.Fields.Item("estable").Value.ToString.PadLeft(3, "0"), writer)
                    createNode("ptoEmi", oRecord.Fields.Item("ptoemi").Value.ToString.PadLeft(3, "0"), writer)
                    createNode("secuencial", oRecord.Fields.Item("secuencial").Value.ToString.PadLeft(9, "0"), writer)
                    createNode("dirMatriz", oRecord.Fields.Item("dirMatriz").Value.ToString, writer)
                    ''Cierre info Tributaria
                    writer.WriteEndElement()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()

                    writer.WriteStartElement("infoFactura")
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery("exec SP_INFO_FACTURA '" & docEntry & "'")
                    createNode("fechaEmision", Date.Parse(oRecord.Fields.Item("DATE").Value.ToString).ToString("dd/MM/yyyy"), writer)
                    createNode("tipoIdentificacionComprador", oRecord.Fields.Item("U_IDENTIFICACION").Value.ToString, writer)
                    createNode("razonSocialComprador", oRecord.Fields.Item("CardName").Value.ToString, writer)
                    createNode("identificacionComprador", oRecord.Fields.Item("U_DOCUMENTO").Value.ToString, writer)
                    createNode("totalSinImpuestos", oRecord.Fields.Item("sin_impuesto").Value.ToString, writer)
                    createNode("totalDescuento", oRecord.Fields.Item("totDescuento").Value.ToString, writer)

                    writer.WriteStartElement("totalConImpuestos")
                    Dim importeTotal = oRecord.Fields.Item("DocTotal").Value.ToString

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()

                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery("exec SP_Total_Con_Impuesto '" & docEntry & "'")
                    If oRecord.RecordCount > 0 Then
                        While oRecord.EoF = False
                            writer.WriteStartElement("totalImpuesto")
                            createNode("codigo", oRecord.Fields.Item(0).Value.ToString, writer)
                            createNode("codigoPorcentaje", oRecord.Fields.Item(1).Value.ToString, writer)
                            createNode("baseImponible", oRecord.Fields.Item(2).Value.ToString, writer)
                            createNode("valor", oRecord.Fields.Item(2).Value.ToString, writer)
                            createNode("baseNoGraIva", Double.Parse("0.00").ToString, writer)
                            createNode("baseImponible", "", writer)
                            writer.WriteEndElement()
                            oRecord.MoveNext()
                        End While
                    End If


                    ''Cierre TotalConImpuestos
                    writer.WriteEndElement()
                    createNode("importeTotal", importeTotal, writer)
                    ''Cierre INFO FACTURA
                    writer.WriteEndElement()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()

                    writer.WriteStartElement("detalles")
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery("exec sp_DetalleFac '" & docEntry & "'")


                    If oRecord.RecordCount > 0 Then

                        While oRecord.EoF = False
                            Dim oRecord2 As SAPbobsCOM.Recordset
                            oRecord2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            writer.WriteStartElement("detalle")
                            createNode("descripcion", oRecord.Fields.Item(0).Value.ToString, writer)
                            createNode("cantidad", oRecord.Fields.Item(1).Value.ToString, writer)
                            createNode("precioUnitario", oRecord.Fields.Item(2).Value.ToString, writer)
                            createNode("descuento", oRecord.Fields.Item(3).Value.ToString, writer)
                            createNode("precioTotalSinImpuesto", oRecord.Fields.Item(4).Value.ToString, writer)
                            writer.WriteStartElement("impuestos")
                            oRecord2.DoQuery("exec SP_Impuesto_Detalle '" & docEntry & "','" & oRecord.Fields.Item(5).Value.ToString & "'")
                            If oRecord2.RecordCount > 0 Then
                                While oRecord2.EoF = False
                                    writer.WriteStartElement("impuesto")
                                    createNode("codigo", oRecord2.Fields.Item(0).Value.ToString, writer)
                                    createNode("codigoPorcentaje", oRecord2.Fields.Item(1).Value.ToString, writer)
                                    createNode("tarifa", oRecord2.Fields.Item(3).Value.ToString, writer)
                                    createNode("baseImponible", oRecord2.Fields.Item(2).Value.ToString, writer)
                                    createNode("valor", oRecord2.Fields.Item(2).Value.ToString, writer)
                                    writer.WriteEndElement()
                                    oRecord2.MoveNext()
                                End While
                            End If

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord2)
                            oRecord2 = Nothing
                            GC.Collect()
                            writer.WriteEndElement()

                            writer.WriteEndElement()
                            oRecord.MoveNext()
                        End While
                    End If

                    ''Cierre detalles
                    writer.WriteEndElement()
                    ''Cierre Factura
                    writer.WriteEndElement()
                    writer.WriteEndDocument()
                    writer.Close()
                    oProgressive.Value += 1
                Next
                oProgressive.Stop()
                oProgressive = Nothing
            End If
        Catch ex As Exception
            Me.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
    End Sub
    Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement(pID)

        writer.WriteString(pName)
        writer.WriteEndElement()

    End Sub

    Private Function claveAcceso(oRecord As SAPbobsCOM.Recordset) As String
        Dim clacceso As String = ""
        Dim fecha = Date.Parse(oRecord.Fields.Item(8).Value).ToString("dd/MM/yyyy")
        clacceso = clacceso & Replace(fecha, "-", "") & "01" & oRecord.Fields.Item(3).Value.ToString & oRecord.Fields.Item(0).Value.ToString & "001001".PadLeft(6, "0") & oRecord.Fields.Item("numcompro").Value.ToString.PadLeft(9, "0") & oRecord.Fields.Item("numcompro").Value.ToString.PadLeft(8, "0")
        clacceso = clacceso & oRecord.Fields.Item(1).Value.ToString & invertirCadena(clacceso)
        'clacceso = clacceso & invertirCadena(clacceso)
        Return clacceso
    End Function

    Private Function invertirCadena(claveacceso As String) As String
        Dim cadenaInvertida As String = ""
        Dim pivote As Integer = 2
        Dim cantidadTotal As Integer = 0

        For i As Integer = claveacceso.Count - 1 To 0 Step -1
            If pivote = 8 Then
                pivote = 2
            End If
            Dim temporal = Integer.Parse(claveacceso.Chars(i))
            temporal *= pivote
            pivote += 1
            cantidadTotal += temporal
        Next
        cantidadTotal = 11 - (cantidadTotal Mod 11)
        cadenaInvertida = cantidadTotal.ToString

        Return cadenaInvertida
    End Function
End Class
