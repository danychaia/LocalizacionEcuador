Imports System.Xml

Public Class generarGRXML
    Public Sub generarXML(DocEntry As String, objectType As String, oCompany As SAPbobsCOM.Company, SBO As SAPbouiCOM.Application)
        Try
            Dim doc As New XmlDocument
            Dim oRecord As SAPbobsCOM.Recordset
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRecord.DoQuery("exec ENCABEZADO_FACTURA '" & DocEntry & "','" & objectType & "'")
            Dim writer As New XmlTextWriter("Comprobante (GR) No." & DocEntry.ToString & ".xml", System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement("guiaRemision")
            writer.WriteAttributeString("id", "comprobante")
            writer.WriteAttributeString("version", "1.1.0")
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
            Dim oContriEspecial = oRecord.Fields.Item("contriespecial").Value
            Dim oObliconta = oRecord.Fields.Item("contaobligado").Value
            ''Cierre info Tributaria
            writer.WriteEndElement()

            writer.WriteStartElement("infoGuiaRemision")
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("exec SP_INFO_FACTURA '" & DocEntry & "','" & objectType & "'")
            createNode("dirEstablecimiento", direccion, writer)
            createNode("dirPartida", direccion, writer)
            createNode("razonSocialTransportista", oRecord.Fields.Item(0).Value, writer)
            createNode("tipoIdentificacionTransportista", oRecord.Fields.Item(1).Value.ToString, writer)
            createNode("rucTransportista", oRecord.Fields.Item(2).Value.ToString, writer)
            createNode("obligadoContabilidad", oObliconta, writer)
            createNode("contribuyenteEspecial", oContriEspecial, writer)

            'createNode("guiaRemision", "", writer)
            createNode("fechaIniTransporte", oRecord.Fields.Item(3).Value.ToString, writer)
            createNode("fechaFinTransporte", oRecord.Fields.Item(4).Value.ToString, writer)
            createNode("placa", oRecord.Fields.Item(5).Value.ToString, writer)
            'cierre infoGuiaRemision
            writer.WriteEndElement()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()

            writer.WriteStartElement("destinatarios")
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("exec SP_DetalleFac '" & DocEntry & "','" & objectType & "'")
            If oRecord.RecordCount > 0 Then
                While oRecord.EoF = False
                    Dim inicial = oRecord.Fields.Item(0).Value
                    Dim final = oRecord.Fields.Item(1).Value
                    Dim tipoDoc = oRecord.Fields.Item(2).Value
                    Dim identifi = oRecord.Fields.Item(3).Value
                    Dim motivo = oRecord.Fields.Item(4).Value
                    Dim direguia = oRecord.Fields.Item(5).Value
                    Dim razon = oRecord.Fields.Item(6).Value
                    Dim oRecord2 As SAPbobsCOM.Recordset
                    oRecord2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    oRecord2.DoQuery("EXEC  FACTURAS_GUIA_REMISION " & DocEntry & ",'" & inicial & "','" & final & "','" & oRecord.Fields.Item(3).Value & "',2")
                    If oRecord2.RecordCount > 0 Then
                        While oRecord2.EoF = False
                            writer.WriteStartElement("destinatario")
                            createNode("identificacionDestinatario", Trim(Right(identifi, Len(identifi) - 2)).ToString(), writer)
                            createNode("razonSocialDestinatario", razon, writer)
                            createNode("dirDestinatario", direguia, writer)
                            createNode("motivoTraslado", motivo, writer)
                            createNode("codDocSustento", tipoDoc, writer)
                            Dim oRecord3 As SAPbobsCOM.Recordset
                            oRecord3 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecord3.DoQuery("EXEC  FACTURAS_GUIA_REMISION " & DocEntry & ",'" & oRecord2.Fields.Item(0).Value & "','" & final & "','" & identifi & "',1")
                            If oRecord3.RecordCount > 0 Then
                                writer.WriteStartElement("detalles")
                                While oRecord3.EoF = False
                                    writer.WriteStartElement("detalle")
                                    createNode("codigoInterno", oRecord3.Fields.Item(1).Value, writer)
                                    createNode("descripcion", oRecord3.Fields.Item(2).Value, writer)
                                    createNode("cantidad", oRecord3.Fields.Item(3).Value, writer)
                                    writer.WriteEndElement()
                                    oRecord3.MoveNext()
                                End While
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord3)
                                oRecord3 = Nothing
                                GC.Collect()
                                'fin detalles
                                writer.WriteEndElement()
                            End If
                            'fin destinatarios
                            writer.WriteEndElement()
                            oRecord2.MoveNext()
                        End While
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord2)
                    oRecord2 = Nothing
                    GC.Collect()

                    oRecord.MoveNext()
                End While
            End If
            'Cierre destinatario
            writer.WriteEndElement()
            ''Cierre Factura
            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Close()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement(pID)
        writer.WriteString(pName)
        writer.WriteEndElement()
    End Sub
End Class
