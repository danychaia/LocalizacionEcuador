Imports System.Xml

Public Class generarATS
    Public Sub generarXML(mes As String, ano As String, objectType As String, oCompany As SAPbobsCOM.Company, SBO As SAPbouiCOM.Application)
        Try
            Dim doc As New XmlDocument
            Dim oRecord As SAPbobsCOM.Recordset
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRecord.DoQuery("exec ATS_Encabezado " & ano & "," & mes)

            Dim writer As New XmlTextWriter("Comprobante (ATS) No." & mes & "-" & ano & ".xml", System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement("iva")
            writer.WriteAttributeString("version", "1.0")
            createNode("TipoIDInformante", oRecord.Fields.Item("TipoIDInformante").Value, writer)
            createNode("IdInformante", oRecord.Fields.Item("IdInformante").Value, writer)
            createNode("razonSocial", oRecord.Fields.Item("razonSocial").Value, writer)
            createNode("numEstabRuc", oRecord.Fields.Item("numEstabRuc").Value, writer)
            createNode("totalVentas", oRecord.Fields.Item("totalVentas").Value, writer)
            createNode("codigoOperativo", oRecord.Fields.Item("codigoOperativo").Value, writer)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()
            writer.WriteStartElement("compras")
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("exec ATS_DetalleCompras " & ano & "," & mes)
            While oRecord.EoF = False
                writer.WriteStartElement("detalleCompras")
                DetalleCompras(oRecord, oCompany, SBOApplication, writer, ano, mes)
                'Detalle Compras 
                oRecord.MoveNext()
                writer.WriteEndElement()
            End While
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()

            'Compras notas de credito 
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("exec ATS_NCDetalleCompras " & ano & "," & mes)
            While oRecord.EoF = False
                writer.WriteStartElement("detalleCompras")
                DetalleCompras(oRecord, oCompany, SBOApplication, writer, ano, mes)
                writer.WriteEndElement()
                oRecord.MoveNext()
            End While
            'Compras 
            writer.WriteEndElement()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()



            writer.WriteStartElement("Ventas")

            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("exec ATS_detalleVentas " & ano & "," & mes)
            While oRecord.EoF = False
                writer.WriteStartElement("detalleVentas")
                createNode("tpIdCliente", oRecord.Fields.Item("tpIdCliente").Value, writer)
                createNode("idCliente", oRecord.Fields.Item("idCliente").Value, writer)

                Dim oRecord2 As SAPbobsCOM.Recordset
                oRecord2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord2.DoQuery("exec ATS_denoCli '" & oRecord.Fields.Item("idCliente").Value & "'")

                createNode("denoCli", oRecord2.Fields.Item("denoCli").Value, writer)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord2)
                oRecord2 = Nothing
                GC.Collect()

                createNode("tipoComprobante", oRecord.Fields.Item("tipoComprobante").Value, writer)
                createNode("tipoEmision", oRecord.Fields.Item("tipoEm").Value, writer)
                createNode("numeroComprobantes", oRecord.Fields.Item("numeroComprobantes").Value, writer)
                createNode("baseNoGraIva", oRecord.Fields.Item("baseNoGraIva").Value, writer)
                createNode("baseImponible", oRecord.Fields.Item("baseImponible").Value, writer)
                createNode("baseImpGrav", oRecord.Fields.Item("baseImpGrav").Value, writer)
                createNode("montoIva", oRecord.Fields.Item("montoIva").Value, writer)

                oRecord2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord2.DoQuery("exec ATS_Compensacion " & ano & "," & mes & ",'" & oRecord.Fields.Item("tipoComprobante").Value & "','" & oRecord.Fields.Item("idCliente").Value & "'")
                If oRecord2.RecordCount > 0 Then
                    writer.WriteStartElement("compensaciones")
                    While oRecord2.EoF = False
                        writer.WriteStartElement("compensacion")
                        createNode("tipoCompe", oRecord2.Fields.Item("tipoCompe").Value, writer)
                        createNode("monto", oRecord2.Fields.Item("monto").Value, writer)
                        writer.WriteEndElement()
                        oRecord2.MoveNext()
                    End While
                    'Cierre Compensaciones
                    writer.WriteEndElement()
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord2)
                oRecord2 = Nothing
                GC.Collect()
                createNode("montoIce", oRecord.Fields.Item("montoIce").Value, writer)
                createNode("valorRetIva", oRecord.Fields.Item("valorRetIva").Value, writer)
                createNode("valorRetRenta", oRecord.Fields.Item("valorRetRenta").Value, writer)

                oRecord2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord2.DoQuery("exec ATS_formasDePago " & ano & "," & mes & ",'" & oRecord.Fields.Item("tipoComprobante").Value & "','" & oRecord.Fields.Item("idCliente").Value & "'")
                writer.WriteStartElement("formasDePago")
                While oRecord2.EoF = False
                    createNode("formaPago", oRecord2.Fields.Item("formaPago").Value, writer)
                End While
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord2)
                oRecord2 = Nothing
                GC.Collect()
                writer.WriteEndElement()
                'Fin detalle ventas 
                writer.WriteEndElement()
                oRecord.MoveNext()
            End While
            'Ciere Ventas 
            writer.WriteEndElement()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()

            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("exec ATS_ventasEstablecimiento " & mes & "," & ano)
            writer.WriteStartElement("ventasEstablecimiento")
            While oRecord.EoF = False
                writer.WriteStartElement("ventaEst")
                createNode("codEstab", oRecord.Fields.Item("codEstab").Value, writer)
                createNode("ventasEstab", oRecord.Fields.Item("ventasEstab").Value, writer)
                createNode("ivaCompe", oRecord.Fields.Item("ivaCompe").Value, writer)
                'Cierre ventasEstablecimiento
                writer.WriteEndElement()
                oRecord.MoveNext()
            End While

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()
            'Cierre ventasEstablecimiento
            writer.WriteEndElement()


            'Inicio exportaciones.
            writer.WriteStartElement("exportaciones")
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("EXEC ATS_detalleExportaciones " & ano & "," & mes)
            While oRecord.EoF = False
                writer.WriteStartElement("detalleExportaciones")
                detalleExportaciones(oRecord, oCompany, SBOApplication, writer, ano, mes)                
                writer.WriteEndElement()
                oRecord.MoveNext()
            End While

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()

            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("EXEC ATS_detalleExportacionesNC " & ano & "," & mes)
            While oRecord.EoF = False
                writer.WriteStartElement("detalleExportaciones")
                detalleExportaciones(oRecord, oCompany, SBOApplication, writer, ano, mes)
                writer.WriteEndElement()
                oRecord.MoveNext()
            End While
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()
            'Fin exportaciones
            writer.WriteEndElement()

            'INICIO ANULADOS
            writer.WriteStartElement("anulados")
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecord.DoQuery("")
            'While oRecord.EoF = False

            'End While

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()
            'FINAL ANULADOS
            writer.WriteEndElement()


            ''Cierre Factura
            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement(pID)
        writer.WriteString(pName)
        writer.WriteEndElement()
    End Sub

    Private Sub DetalleCompras(oRecord As SAPbobsCOM.Recordset, oCompany As SAPbobsCOM.Company, application As SAPbouiCOM.Application, writer As XmlTextWriter, ano As String, mes As String)
        createNode("codSustento", oRecord.Fields.Item("CodSustento").Value, writer)
        createNode("tpldProv", oRecord.Fields.Item("tpldProv").Value, writer)
        createNode("idProv", oRecord.Fields.Item("idProv").Value, writer)
        createNode("tipoComprobante", oRecord.Fields.Item("tipoComprobante").Value, writer)
        createNode("parteRel", oRecord.Fields.Item("parteRel").Value, writer)
        createNode("fechaRegistro", oRecord.Fields.Item("fechaRegistro").Value, writer)
        createNode("establecimiento", oRecord.Fields.Item("establecimiento").Value, writer)
        createNode("puntoEmision", oRecord.Fields.Item("puntoEmision").Value, writer)
        createNode("secuencial", oRecord.Fields.Item("secuencial").Value, writer)
        createNode("fechaEmision", oRecord.Fields.Item("fechaEmision").Value, writer)
        createNode("autorizacion", oRecord.Fields.Item("autorizacion").Value, writer)
        createNode("baseNoGraIva", oRecord.Fields.Item("baseNoGraIva").Value, writer)
        createNode("baseImponible", oRecord.Fields.Item("baseImponible").Value, writer)
        createNode("baseImpGrav", oRecord.Fields.Item("baseImpGrav").Value, writer)
        createNode("baseImpExe", oRecord.Fields.Item("baseImpExe").Value, writer)
        createNode("montoIce", oRecord.Fields.Item("montoIce").Value, writer)
        createNode("montoIva", oRecord.Fields.Item("montoIva").Value, writer)
        createNode("valRetBien10", oRecord.Fields.Item("valRetBien10").Value, writer)
        createNode("valRetServ20", oRecord.Fields.Item("valRetServ20").Value, writer)
        createNode("valorRetBienes", oRecord.Fields.Item("valRetServ20").Value, writer)
        createNode("ValorRetServicios", oRecord.Fields.Item("ValorRetServicios").Value, writer)
        'createNode("valRetServ50", oRecord.Fields.Item("ValorRetServicios").Value, writer)
        createNode("valRetServ100", oRecord.Fields.Item("valRetServ100").Value, writer)
        createNode("totbasesImpReemb", oRecord.Fields.Item("totbasesImpReemb").Value, writer)

        writer.WriteStartElement("pagoExterior")
        createNode("pagoLocExt", "", writer)
        createNode("paisEfecPago", "", writer)
        createNode("aplicConvDobTrib", "", writer)
        createNode("pagExtSujRetNorLeg", "", writer)
        'Fin pago exterior
        writer.WriteEndElement()
        'EMPIEZA FORMAS DE PAGO PARA COMPRAS
        Dim oRecordP As SAPbobsCOM.Recordset
        oRecordP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordP.DoQuery("EXEC ATS_formasDePago " & ano & "," & mes & ",'" & oRecord.Fields.Item("tipoComprobante").Value & "','" & oRecord.Fields.Item("tipoComprobante").Value & "'")
        writer.WriteStartElement("formasDePago")
        While oRecordP.EoF
            createNode("formaPago", oRecordP.Fields.Item("formaPago").Value, writer)
            oRecordP.MoveNext()
        End While
        'final forma de pago 
        writer.WriteEndElement()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordP)
        oRecordP = Nothing
        GC.Collect()


        'DETALLE AIR PARA COMPRAS 
        oRecordP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordP.DoQuery("EXEC ATS_Air " & oRecord.Fields.Item("DocEntry").Value)

        While oRecordP.RecordCount
            writer.WriteStartElement("air")
            writer.WriteStartElement("detalleAir")
            createNode("codRetAir", oRecordP.Fields.Item("codRetAir").Value, writer)
            createNode("baseImpAir", oRecordP.Fields.Item("baseImpAir").Value, writer)
            createNode("porcentajeAir", oRecordP.Fields.Item("porcentajeAir").Value, writer)
            createNode("valRetAir", oRecordP.Fields.Item("valRetAir").Value, writer)
            'Fin detalle Air
            writer.WriteEndElement()
        End While
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordP)
        oRecordP = Nothing
        GC.Collect()
        createNode("estabRetencion1", oRecord.Fields.Item("estabRetencion1").Value, writer)
        createNode("ptoEmiRetencion1", oRecord.Fields.Item("ptoEmiRetencion1").Value, writer)
        createNode("secRetencion1", oRecord.Fields.Item("secRetencion1").Value, writer)
        createNode("autRetencion1", oRecord.Fields.Item("autRetencion1").Value, writer)
        createNode("fechaEmiRet1", oRecord.Fields.Item("fechaEmiRet1").Value, writer)
    End Sub

    Private Sub detalleExportaciones(oRecord As SAPbobsCOM.Recordset, oCompany As SAPbobsCOM.Company, application As SAPbouiCOM.Application, writer As XmlTextWriter, ano As String, mes As String)
        createNode("tpIdClienteEx", oRecord.Fields.Item("tpIdClienteEx").Value, writer)
        createNode("idClienteEx", oRecord.Fields.Item("idClienteEx").Value, writer)
        createNode("parteRelExp", oRecord.Fields.Item("parteRelExp").Value, writer)
        createNode("tipoRegi", oRecord.Fields.Item("tipoRegi").Value, writer)
        createNode("paisEfecPagoGen", oRecord.Fields.Item("paisEfecPagoGen").Value, writer)
        createNode("paisEfecExp", oRecord.Fields.Item("paisEfecExp").Value, writer)
        createNode("exportacionDe", oRecord.Fields.Item("exportacionDe").Value, writer)
        createNode("tipoComprobante", oRecord.Fields.Item("tipoComprobante").Value, writer)
        createNode("distAduanero", oRecord.Fields.Item("distAduanero").Value, writer)
        createNode("anio", oRecord.Fields.Item("anio").Value, writer)
        createNode("regimen", oRecord.Fields.Item("regimen").Value, writer)
        createNode("docTransp", oRecord.Fields.Item("docTransp").Value, writer)
        createNode("fechaEmbarque", oRecord.Fields.Item("fechaEmbarque").Value, writer)
        createNode("valorFOB", oRecord.Fields.Item("valorFOB").Value, writer)
        createNode("valorFOBComprobante", oRecord.Fields.Item("valorFOBComprobante").Value, writer)
        createNode("establecimiento", oRecord.Fields.Item("establecimiento").Value, writer)
        createNode("puntoEmision", oRecord.Fields.Item("puntoEmision").Value, writer)
        createNode("autorizacion", oRecord.Fields.Item("autorizacion").Value, writer)
        createNode("fechaEmision", oRecord.Fields.Item("fechaEmision").Value, writer)

    End Sub

End Class
