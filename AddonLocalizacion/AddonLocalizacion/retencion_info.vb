Public Class retencion_info
    Private XmlForm As String = Replace(Application.StartupPath & "\retencion_info.srf", "\\", "\")
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter
    Private selected As Boolean = False
    Public Sub New()
        MyBase.New()
        Try
            Me.SBO_Application = UDT_UF.SBOApplication
            Me.oCompany = UDT_UF.Company

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "rInf") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("rInf")
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
                inicio = oForm.Items.Item("Item_1").Specific
                fin = oForm.Items.Item("Item_3").Specific
                inicio.DataBind.SetBound(True, "", "Date")
                fin.DataBind.SetBound(True, "", "Date2")

            Else
                oForm = Me.SBO_Application.Forms.Item("rInf")
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
        If pVal.FormTypeEx = "60006" And pVal.Before_Action = True And pVal.FormUID = "rInf" Then
            
            If pVal.ItemUID = "Item_4" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim oDe As SAPbouiCOM.EditText
                Dim oHasta As SAPbouiCOM.EditText
                oDe = oForm.Items.Item("Item_1").Specific
                oHasta = oForm.Items.Item("Item_3").Specific
                If oDe.Value = "" Or oHasta.Value = "" Then
                    SBOApplication.SetStatusBarMessage("Debe de seleccionar un rango de fecha", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                Else
                    visualizardata(oDe.Value, oHasta.Value)
                End If
            End If
            If pVal.ItemUID = "Item_6" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim gridView As SAPbouiCOM.Grid
                gridView = oForm.Items.Item("Item_6").Specific
                If pVal.Row <> -1 Then
                    Me.selected = True
                    Dim docentry = gridView.DataTable.GetValue("DocEntry", pVal.Row).ToString
                    UDT_UF.docEntry = docentry
                    Dim detalle As New retencion_info_detalle
                    BubbleEvent = False
                End If
            End If
            If pVal.ItemUID = "Item_5" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                If selected Then
                    Try
                        Dim lRetCode As Integer = -1
                        Dim MensajeError As String = ""
                        ' oCompany.StartTransaction()
                        Dim oInvoiceAnula As SAPbobsCOM.Documents
                        Dim oCan As SAPbobsCOM.Documents
                        Dim oDe As SAPbouiCOM.EditText
                        Dim oHasta As SAPbouiCOM.EditText
                        Dim oRecored As SAPbobsCOM.Recordset
                        oDe = oForm.Items.Item("Item_1").Specific
                        oHasta = oForm.Items.Item("Item_3").Specific
                        oInvoiceAnula = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                        oRecored = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecored.DoQuery("UPDATE OPCH SET U_ESTADO = 'C' WHERE DocEntry =" & UDT_UF.docEntry)
                        selected = False
                        'obtenendo el numero de la factura
                        If oInvoiceAnula.GetByKey(UDT_UF.docEntry) Then
                            'oCan = oInvoiceAnula.CreateCancellationDocument
                            ' oCan.CreateCancellationDocument()
                            '   lRetCode = oCan.Add()
                            'If lRetCode <> 0 Then
                            'MensajeError = String.Format("Anula Factura {0}-{1}", oCompany.GetLastErrorCode, oCompany.GetLastErrorDescription)
                            '  SBO_Application.SetStatusBarMessage(MensajeError, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            ' oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            ' Else
                            ' oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            ' End If
                        Else
                        'oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                       

                        visualizardata(oDe.Value, oHasta.Value)
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    End Try
                   
                End If
            End If
        End If
    End Sub

    Private Sub visualizardata(p1 As String, p2 As String)
        Try
            Dim gridView As SAPbouiCOM.Grid
            gridView = oForm.Items.Item("Item_6").Specific
            Dim sql As String = "EXEC BUSCAR_INFO_RETENCION '" & p1 & "','" & p2 & "'"
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(sql)
            gridView.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            gridView.AutoResizeColumns()
            gridView.Columns.Item(0).Editable = False
            Dim oCol As SAPbouiCOM.GridColumn

            oCol = gridView.Columns.Item(0)

            oCol.LinkedObjectType = 18
            gridView.Columns.Item(1).Editable = False
            gridView.Columns.Item(2).Editable = False
            gridView.Columns.Item(3).Editable = False
            gridView.Columns.Item(4).Editable = False

            System.Runtime.InteropServices.Marshal.ReleaseComObject(gridView)
            gridView = Nothing
            GC.Collect()
        Catch ex As Exception
            Me.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
        End Try
    End Sub

End Class
