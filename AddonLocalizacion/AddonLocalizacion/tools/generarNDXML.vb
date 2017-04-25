Imports System.Xml

Public Class generarNDXML
    Public Sub generarXML(DocEntry As String, objectType As String, oCompany As SAPbobsCOM.Company, SBO As SAPbouiCOM.Application)
        Dim doc As New XmlDocument
        Dim oRecord As SAPbobsCOM.Recordset
        Dim oContriEspecial As String
        Dim oObliconta As String
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecord.DoQuery("exec ENCABEZADO_FACTURA '" & DocEntry & "','ND'")
        Dim writer As New XmlTextWriter("Comprobante (ND) No." & DocEntry.ToString & ".xml", System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("notaDebito")
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
        oContriEspecial = oRecord.Fields.Item("contriespecial").Value
        oObliconta = oRecord.Fields.Item("contaobligado").Value
        ''Cierre info Tributaria
        writer.WriteEndElement()

        writer.WriteStartElement("infoNotaDebito")
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecord.DoQuery("exec SP_INFO_FACTURA '" & DocEntry & "','ND'")
        createNode("fechaEmision", Date.Parse(oRecord.Fields.Item("DATE").Value.ToString).ToString("dd/MM/yyyy"), writer)
        createNode("tipoIdentificacionComprador", oRecord.Fields.Item("U_IDENTIFICACION").Value.ToString, writer)
        createNode("razonSocialComprador", oRecord.Fields.Item("CardName").Value.ToString, writer)
        createNode("identificacionComprador", oRecord.Fields.Item("U_DOCUMENTO").Value.ToString, writer)
        createNode("contribuyenteEspecial", oContriEspecial, writer)
        createNode("obligadoContabilidad", oObliconta, writer)
        createNode("codDocModificado", "01", writer)
        createNode("numDocModificado", oRecord.Fields.Item("DocModifi").Value, writer)
        createNode("fechaEmisionDocSustento", oRecord.Fields.Item("FechaModifi").Value, writer)
        createNode("totalSinImpuestos", oRecord.Fields.Item("sin_impuesto").Value.ToString, writer)
       
        Dim importeTotal = oRecord.Fields.Item("DocTotal").Value.ToString
        Dim moneda = oRecord.Fields.Item("MONEDA").Value.ToString
        Dim motivo = oRecord.Fields.Item("Comments").Value.ToString
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
        oRecord = Nothing
        GC.Collect()
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecord.DoQuery("exec SP_Impuesto_Detalle '" & DocEntry & "','','ND'")
        If oRecord.RecordCount > 0 Then
            writer.WriteStartElement("impuestos")
            While oRecord.EoF = False
                writer.WriteStartElement("impuesto")
                createNode("codigo", oRecord.Fields.Item(0).Value.ToString, writer)
                createNode("codigoPorcentaje", oRecord.Fields.Item(1).Value.ToString, writer)
                createNode("tarifa", oRecord.Fields.Item(3).Value.ToString, writer)
                createNode("baseImponible", oRecord.Fields.Item(2).Value.ToString, writer)
                createNode("valor", oRecord.Fields.Item(4).Value.ToString, writer)
                writer.WriteEndElement()
                oRecord.MoveNext()
            End While
            ''Cierre de impuestos 
            writer.WriteEndElement()
        End If
        createNode("valorTotal", importeTotal, writer)
        

        ''Cierre infoNotaDebito
        writer.WriteEndElement()
        writer.WriteStartElement("motivos")
        writer.WriteStartElement("motivo")
        createNode("razon", motivo, writer)
        createNode("valor", importeTotal, writer)
        writer.WriteEndElement()
        writer.WriteEndElement()
        ''Cierre Nota de Débito
        writer.WriteEndElement()
        writer.Close()

    End Sub

    Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement(pID)
        writer.WriteString(pName)
        writer.WriteEndElement()
    End Sub
End Class
