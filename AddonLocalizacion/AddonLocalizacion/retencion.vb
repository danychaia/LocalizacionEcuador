Imports System.Xml
Imports System.Globalization

Public Class retencion
    Private XmlForm As String = Replace(Application.StartupPath & "\retencion_comprobante.srf", "\\", "\")
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

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "fRtn") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("fRtn")
                oForm.Left = 400
                oForm.Visible = True
                oForm.PaneLevel = 1

                Dim inicio As SAPbouiCOM.EditText
                Dim fin As SAPbouiCOM.EditText
                Dim cmdenviar As SAPbouiCOM.Button
                'esto es para poder hacer que los textos tengan formato de fecha
                oForm.DataSources.DataTables.Add("MyDataTable")
                oForm.DataSources.UserDataSources.Add("Date", SAPbouiCOM.BoDataType.dt_DATE)
                oForm.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_DATE)
                inicio = oForm.Items.Item("Item_0").Specific
                fin = oForm.Items.Item("Item_1").Specific
                inicio.DataBind.SetBound(True, "", "Date")
                fin.DataBind.SetBound(True, "", "Date2")

            Else
                oForm = Me.SBO_Application.Forms.Item("fRtn")
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

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        If pVal.FormTypeEx = "60006" And pVal.Before_Action = True Then
            If pVal.ItemUID = "btn1" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim oDe As SAPbouiCOM.EditText
                Dim oHasta As SAPbouiCOM.EditText
                oDe = oForm.Items.Item("Item_0").Specific
                oHasta = oForm.Items.Item("Item_1").Specific
                If oDe.Value = "" Or oHasta.Value = "" Then
                    SBOApplication.SetStatusBarMessage("Debe de seleccionar un rango de fecha", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                Else
                    generaRetencionXML(oDe.Value, oHasta.Value.ToString, SBOApplication)
                End If
            End If
        End If
    End Sub

    Private Sub generaRetencionXML(p1 As String, p2 As String, app As SAPbouiCOM.Application)
        Dim doc As New XmlDocument
        Dim oRecord As SAPbobsCOM.Recordset
        Dim oRecordU As SAPbobsCOM.Recordset
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim writer As New XmlTextWriter("Comprobante (RETENCION) No.1.xml", System.Text.Encoding.UTF8)
        Dim oProgressive As SAPbouiCOM.ProgressBar

        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("iva")
        oRecord.DoQuery(" EXEC SP_IDENTIFICACION_INFORMANTE '" & oCompany.CompanyName & "'")
        If oRecord.RecordCount > 0 Then
            createNode("TipoIDInformante", oRecord.Fields.Item(0).Value.ToString, writer)
            createNode("IdInformante", oRecord.Fields.Item(1).Value.ToString, writer)
            createNode("razonSocial", oRecord.Fields.Item(2).Value.ToString, writer)
            createNode("Anio", p1.Substring(0, 4), writer)
            createNode("Mes", p2.Substring(4, 2), writer)
            createNode("numEstabRuc", oRecord.Fields.Item(3).Value.ToString, writer)
            createNode("totalVentas", "", writer)
            createNode("codigoOperativo", "IVA", writer)
        End If

        writer.WriteStartElement("compras")
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
        oRecord = Nothing
        GC.Collect()

        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecord.DoQuery("SP_COMPRA_DETALLE_RETENCION '14604'")
        If oRecord.RecordCount > 0 Then
            oProgressive = SBO_Application.StatusBar.CreateProgressBar("Generando Retencion de :", oRecord.RecordCount, True)
            While oRecord.EoF = False
                writer.WriteStartElement("detalleCompras")
                createNode("codSustento", oRecord.Fields.Item(1).Value, writer)
                createNode("tpIdProv", oRecord.Fields.Item(2).Value, writer)
                createNode("idProv", oRecord.Fields.Item(3).Value, writer)
                createNode("tipoComprobante", oRecord.Fields.Item(4).Value, writer)
                createNode("parteRel", oRecord.Fields.Item(6).Value, writer)
                createNode("fechaRegistro", oRecord.Fields.Item(5).Value, writer)
                createNode("establecimiento", oRecord.Fields.Item(7).Value.ToString.Substring(0, 3), writer)
                createNode("puntoEmision", oRecord.Fields.Item(7).Value.ToString.Substring(3, 3), writer)
                createNode("secuencial", oRecord.Fields.Item(7).Value.ToString.Substring(6, 7), writer)
                createNode("fechaEmision", oRecord.Fields.Item(5).Value, writer)
                createNode("autorizacion", "", writer)

                Dim oRecord2 As SAPbobsCOM.Recordset
                oRecord2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord2.DoQuery("exec SP_RETENCION_SUMATORIAS " & oRecord.Fields.Item(0).Value & ",'parther'")
                If oRecord2.RecordCount > 0 Then
                    While oRecord2.EoF = False
                        createNode("baseNoGraIva", Double.Parse(oRecord2.Fields.Item(0).Value), writer)
                        createNode("baseImponible", Double.Parse(oRecord2.Fields.Item(1).Value), writer)
                        createNode("baseImpGrav", Double.Parse(oRecord2.Fields.Item(2).Value), writer)
                        createNode("baseImpExe", Double.Parse(oRecord2.Fields.Item(3).Value), writer)
                        createNode("baseImpExe", Double.Parse(oRecord2.Fields.Item(3).Value), writer)
                        createNode("montoIva", Double.Parse(oRecord2.Fields.Item(4).Value), writer)
                        createNode("montoIce", Double.Parse(oRecord2.Fields.Item(5).Value), writer)
                        createNode("valRetBien10", "0.00", writer)
                        createNode("valRetServ20", "0.00", writer)
                        createNode("valorRetBienes", "0.00", writer)
                        createNode("valRetServ50", "0.00", writer)
                        createNode("valorRetServicios", "0.00", writer)
                        createNode("valRetServ100", "0.00", writer)
                        createNode("totbasesImpReemb", "0.00", writer)                      
                        oRecord2.MoveNext()
                    End While
                End If
                writer.WriteStartElement("pagoExterior")
                createNode("pagoLocExt", oRecord.Fields.Item(8).Value, writer)
                createNode("paisEfecPago", "NA", writer)
                createNode("aplicConvDobTrib", "NA", writer)
                createNode("pagExtSujRetNorLeg", "NA", writer)
                writer.WriteEndElement()





                Dim oRecord3 As SAPbobsCOM.Recordset
                oRecord3 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord3.DoQuery("exec SP_COMPRA_DETALLE_RETENCION_AIR '" & oRecord.Fields.Item(0).Value & "'")
                If oRecord3.RecordCount > 0 Then
                    While oRecord3.EoF = False
                        writer.WriteStartElement("detalleAir")
                        createNode("codRetAir", oRecord3.Fields.Item(0).Value, writer)
                        createNode("baseImpAir", oRecord3.Fields.Item(1).Value, writer)
                        createNode("porcentajeAir", oRecord3.Fields.Item(2).Value, writer)
                        createNode("valRetAir", oRecord3.Fields.Item(3).Value, writer)
                        writer.WriteEndElement()
                        oRecord3.MoveNext()
                    End While
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord3)
                oRecord3 = Nothing
                GC.Collect()

                createNode("estabRetencion1", oRecord.Fields.Item(10).Value.ToString.Substring(0, 3), writer)
                createNode("ptoEmiRetencion1", oRecord.Fields.Item(10).Value.ToString.Substring(3, 3), writer)
                createNode("secRetencion1", oRecord.Fields.Item(10).Value.ToString.Substring(6, 7), writer)
                createNode("autRetencion1", oRecord.Fields.Item(11).Value, writer)
                createNode("fechaEmiRet1", oRecord.Fields.Item(5).Value, writer)

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord2)
                oRecord2 = Nothing
                GC.Collect()

                writer.WriteEndElement()
                oRecordU = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordU.DoQuery("UPDATE OPCH SET U_ESTADO='G' WHERE DocEntry=" & oRecord.Fields.Item(0).Value)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordU)
                oRecordU = Nothing
                GC.Collect()
                oRecord.MoveNext()
                oProgressive.Value += 1
            End While
            oProgressive.Stop()
            oProgressive = Nothing
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
        oRecord = Nothing
        GC.Collect()


        'FINALIZA TAG COMPRAS
        writer.WriteEndElement()
        'FINALIZA TAG iva
        writer.WriteEndElement()
        writer.WriteEndDocument()
        writer.Close()

    End Sub

    Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement(pID)

        writer.WriteString(pName)
        writer.WriteEndElement()
    End Sub
End Class
