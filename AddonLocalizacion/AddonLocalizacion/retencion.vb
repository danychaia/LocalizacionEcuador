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
                    generaRetencionXML(oDe.Value, oHasta.Value, SBOApplication)
                End If
            End If
        End If
    End Sub

    Private Sub generaRetencionXML(p1 As String, p2 As String, app As SAPbouiCOM.Application)
        Dim doc As New XmlDocument
        Dim oRecord As SAPbobsCOM.Recordset
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim writer As New XmlTextWriter("Comprobante (RETENCION) No.1.xml", System.Text.Encoding.UTF8)

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
        oRecord.DoQuery("SP_COMPRA_DETALLE_RETENCION '" & p1 & "','" & p2 & "'")
        If oRecord.RecordCount > 0 Then
            While oRecord.EoF = False
                writer.WriteStartElement("detalleCompras")
                createNode("codSustento", oRecord.Fields.Item(1).Value, writer)
                createNode("tpIdProv", oRecord.Fields.Item(2).Value, writer)
                createNode("idProv", oRecord.Fields.Item(3).Value, writer)
                createNode("tipoComprobante", oRecord.Fields.Item(4).Value, writer)
                createNode("parteRel", "NO", writer)
                createNode("fechaRegistro", oRecord.Fields.Item(5).Value, writer)
                createNode("establecimiento", oRecord.Fields.Item(6).Value, writer)
                createNode("puntoEmision", oRecord.Fields.Item(7).Value, writer)
                createNode("secuencial", oRecord.Fields.Item(8).Value, writer)
                createNode("fechaEmision", oRecord.Fields.Item(5).Value, writer)
                createNode("autorizacion", oRecord.Fields.Item(9).Value, writer)

                writer.WriteEndElement()
                oRecord.MoveNext()
            End While
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
