Imports System.Xml

Public Class generarRetencionXML
    Public Sub generaRetencionXML(docEntry As String, Tipo As String, ByVal SBO_Application As SAPbouiCOM.Application, ByVal oCompany As SAPbobsCOM.Company)
        Dim doc As New XmlDocument
        Dim oRecord As SAPbobsCOM.Recordset
        Dim oRecordU As SAPbobsCOM.Recordset
        Dim ruta = Application.StartupPath & "\" & Date.Now.Year.ToString & "-" & Date.Now.Month.ToString & ".xml"
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try

            '  If My.Computer.FileSystem.FileExists(ruta) = False Then

            Dim writer As New XmlTextWriter(Date.Now.Year.ToString & "-" & Date.Now.Month & "-" & docEntry.ToString & ".xml", System.Text.Encoding.UTF8)
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
                createNode("Anio", Date.Now.Year, writer)
                createNode("Mes", Date.Now.Month, writer)
                createNode("numEstabRuc", oRecord.Fields.Item(3).Value.ToString, writer)
                createNode("totalVentas", "", writer)
                createNode("codigoOperativo", "IVA", writer)
            End If

            writer.WriteStartElement("compras")
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()

            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("SP_COMPRA_DETALLE_RETENCION '" & docEntry & "'")
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
            ' Else
            ' AgregarNodo(docEntry, Tipo, SBO_Application, oCompany, ruta)
            '   End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub
    Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement(pID)

        writer.WriteString(pName)
        writer.WriteEndElement()
    End Sub

  

    Private Sub AgregarNodo(docEntry As String, Tipo As String, SBO_Application As SAPbouiCOM.Application, oCompany As SAPbobsCOM.Company, url As String)
        Try
            Dim doc As New XmlDocument
            doc.Load(url)
            Dim node = doc.DocumentElement.SelectSingleNode("/iva/compras")
            Dim detalle = doc.CreateElement("detalleCompras")
            Dim codSustento = doc.CreateElement("codSustento", "1")
            detalle.AppendChild(codSustento)
            node.AppendChild(detalle)
            doc.LoadXml(node.InnerXml)
            doc.Save(url)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub generaXML(DocEntry As String, objectType As String, oCompany As SAPbobsCOM.Company, SBO As SAPbouiCOM.Application)
        Try
            Dim doc As New XmlDocument
            Dim oRecord As SAPbobsCOM.Recordset
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("exec ENCABEZADO_FACTURA '" & DocEntry & "','RTNC'")
            Dim writer As New XmlTextWriter("Comprobante (RC) No." & DocEntry.ToString & ".xml", System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement("comprobanteRetencion")
            writer.WriteAttributeString("id", "comprobante")
            writer.WriteAttributeString("version", "1.0.0")
            writer.WriteStartElement("infoTributaria")
            createNode("razonSocial", oRecord.Fields.Item(2).Value.ToString, writer)
            'createNode("ambiente", oRecord.Fields.Item(0).Value.ToString, writer)
            'createNode("tipoEmision", oRecord.Fields.Item(1).Value.ToString, writer)
            createNode("ruc", oRecord.Fields.Item(3).Value.ToString.PadLeft(13, "0"), writer)
            'createNode("claveAcesso", claveAcceso(oRecord).PadLeft(49, "0"), writer)
            'createNode("claveAcesso", "", writer)
            createNode("codDoc", oRecord.Fields.Item("codDoc").Value.ToString.PadLeft(2, "0"), writer)
            createNode("estab", oRecord.Fields.Item("estable").Value.ToString.PadLeft(3, "0"), writer)
            createNode("ptoEmi", oRecord.Fields.Item("ptoemi").Value.ToString.PadLeft(3, "0"), writer)
            createNode("secuencial", oRecord.Fields.Item("secuencial").Value.ToString.PadLeft(9, "0"), writer)
            createNode("dirMatriz", oRecord.Fields.Item("dirMatriz").Value.ToString, writer)
            Dim direccion = oRecord.Fields.Item("dirMatriz").Value.ToString
            ''Cierre info Tributaria
            writer.WriteEndElement()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()



            writer.WriteStartElement("infoCompRetencion")
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("exec SP_INFO_FACTURA '" & DocEntry & "','RTNC'")
            createNode("fechaEmision", Date.Parse(oRecord.Fields.Item("DocDate").Value.ToString).ToString("dd/MM/yyyy"), writer)
            createNode("dirEstablecimiento", oRecord.Fields.Item(1).Value, writer)
            createNode("contribuyenteEspecial", oRecord.Fields.Item(2).Value, writer)
            createNode("obligadoContabilidad", oRecord.Fields.Item(3).Value, writer)
            createNode("tipoIdentificacionSujetoRetenido", oRecord.Fields.Item("U_IDENTIFICACION").Value.ToString, writer)
            createNode("razonSocialSujetoRetenido", oRecord.Fields.Item("CardName").Value.ToString, writer)
            createNode("identificacionSujetoRetenido", oRecord.Fields.Item("U_DOCUMENTO").Value.ToString, writer)
            createNode("periodoFiscal", oRecord.Fields.Item("YEAR").Value.ToString() & "/" & oRecord.Fields.Item("MONTH").Value.ToString(), writer)
            ''Cierre infoCompRetencion
            writer.WriteEndElement()
                      
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()



            ''Cierre ComprobanteRetencion
            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

End Class
