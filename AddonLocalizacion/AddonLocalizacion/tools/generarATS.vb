Imports System.Xml

Public Class generarATS
    Public Sub generarXML(DocEntry As String, objectType As String, oCompany As SAPbobsCOM.Company, SBO As SAPbouiCOM.Application)
        Try
            Dim doc As New XmlDocument
            Dim oRecord As SAPbobsCOM.Recordset
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRecord.DoQuery("")
            Dim writer As New XmlTextWriter("Comprobante (ATS) No." & Date.Now.Month & ".xml", System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement("iva")
            writer.WriteAttributeString("version", "1.0")
            createNode("TipoIDInformante", "", writer)
            createNode("IdInformante", "", writer)
            createNode("razonSocial", "", writer)
            createNode("numEstabRuc", "", writer)
            createNode("totalVentas", "", writer)
            createNode("codigoOperativo", "", writer)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()
            writer.WriteStartElement("compras")
            writer.WriteStartElement("detalleCompras")
            createNode("codSustento", "", writer)
            createNode("idProv", "", writer)
            createNode("parteRel", "", writer)
            createNode("fechaRegistro", "", writer)
            createNode("establecimiento", "", writer)
            createNode("puntoEmision", "", writer)
            createNode("secuencial", "", writer)
            createNode("fechaEmision", "", writer)
            createNode("autorizacion", "", writer)
            createNode("baseNoGraIva", "", writer)
            createNode("baseImponible", "", writer)
            createNode("baseImpGrav", "", writer)
            createNode("baseImpExe", "", writer)
            createNode("montoIce", "", writer)
            createNode("montoIva", "", writer)
            createNode("valRetBien10", "", writer)
            createNode("valRetServ20", "", writer)
            createNode("valorRetBienes", "", writer)
            createNode("valRetBien10", "", writer)
            createNode("valRetServ50", "", writer)
            createNode("valorRetServicios", "", writer)
            createNode("valRetServ100", "", writer)
            createNode("totbasesImpReemb", "", writer)

            writer.WriteStartElement("pagoExterior")
            createNode("pagoLocExt", "", writer)
            createNode("paisEfecPago", "", writer)
            createNode("aplicConvDobTrib", "", writer)
            createNode("pagExtSujRetNorLeg", "", writer)
            'Fin pago exterior
            writer.WriteEndElement()

            writer.WriteStartElement("air")

            writer.WriteStartElement("detalleAir")
            createNode("codRetAir", "", writer)
            createNode("baseImpAir", "", writer)
            createNode("porcentajeAir", "", writer)
            createNode("valRetAir", "", writer)
            'Fin detalle Air
            writer.WriteEndElement()
            'Fin air
            writer.WriteEndElement()

            createNode("estabRetencion1", "", writer)
            createNode("ptoEmiRetencion1", "", writer)
            createNode("secRetencion1", "", writer)
            createNode("autRetencion1", "", writer)
            createNode("fechaEmiRet1", "", writer)
            'Detalle Compras 
            writer.WriteEndElement()
            'Compras 
            writer.WriteEndElement()

            writer.WriteStartElement("Ventas")
            writer.WriteStartElement("detalleVentas")
            createNode("tpIdCliente", "", writer)
            createNode("idCliente", "", writer)
            createNode("tipoComprobante", "", writer)
            createNode("tipoEmision", "", writer)
            createNode("numeroComprobantes", "", writer)
            createNode("baseNoGraIva", "", writer)
            createNode("baseImponible", "", writer)
            createNode("baseImpGrav", "", writer)
            createNode("montoIva", "", writer)
            createNode("montoIce", "", writer)
            createNode("valorRetIva", "", writer)
            createNode("valorRetRenta", "", writer)
            'Fin detalle ventas 
            writer.WriteEndElement()

            'Ciere Ventas 
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
End Class
