Public Class retencion_info_ventas
    Private XmlForm As String = Replace(Application.StartupPath & "\retencion_info_Ventas.srf", "\\", "\")
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

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "rvin") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("rvin")
                oForm.Left = 400
                oForm.Visible = True
                oForm.PaneLevel = 1
                oForm.DataSources.DataTables.Add("MyDataTable")

                Dim inicio As SAPbouiCOM.EditText
                Dim fin As SAPbouiCOM.EditText
                Dim cmdenviar As SAPbouiCOM.Button
                'esto es para poder hacer que los textos tengan formato de fecha
                oForm.DataSources.UserDataSources.Add("Date", SAPbouiCOM.BoDataType.dt_DATE)
                oForm.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_DATE)
                inicio = oForm.Items.Item("Item_1").Specific
                fin = oForm.Items.Item("Item_3").Specific
                inicio.DataBind.SetBound(True, "", "Date")
                fin.DataBind.SetBound(True, "", "Date2")


            Else
                oForm = Me.SBO_Application.Forms.Item("rvin")
            End If
            'visualizardata()
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

    Private Sub visualizardata(p1 As String, p2 As String)
        Try
            Dim gridView As SAPbouiCOM.Grid
            gridView = oForm.Items.Item("Item_6").Specific
            Dim sql As String = "BUSCAR_FACTURA_FACE '" & p1 & "','" & p2 & "'"
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(sql)
            gridView.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            gridView.AutoResizeColumns()
            gridView.Columns.Item(0).Editable = False
            Dim oCol As SAPbouiCOM.GridColumn

            oCol = gridView.Columns.Item(0)

            oCol.LinkedObjectType = 13
            gridView.Columns.Item(1).Editable = False
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


    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormTypeEx = "60006" And pVal.Before_Action = True And pVal.FormUID = "rvin" Then
                If pVal.ItemUID = "Item_4" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                    Dim oDe As SAPbouiCOM.EditText
                    Dim oHasta As SAPbouiCOM.EditText
                    oDe = oForm.Items.Item("Item_1").Specific
                    oHasta = oForm.Items.Item("Item_3").Specific
                    If oDe.Value = "" Or oHasta.Value = "" Then
                        SBOApplication.SetStatusBarMessage("Debe de seleccionar un rango de fecha", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    Else
                        visualizardata(oDe.Value, oHasta.Value)
                        BubbleEvent = False
                        Return
                    End If
                End If
            End If
            If pVal.ItemUID = "Item_6" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim gridView As SAPbouiCOM.Grid
                gridView = oForm.Items.Item("Item_6").Specific
                Dim detalle As New retencion_info_ventas_detalle(gridView.DataTable.GetValue(gridView.DataTable.Columns.Item(0).Name, pVal.Row), gridView.DataTable.GetValue(gridView.DataTable.Columns.Item(1).Name, pVal.Row))
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
      
    End Sub
End Class
