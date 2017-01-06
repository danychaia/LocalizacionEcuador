Public Class comprobantes_autorizados
    Private XmlForm As String = Replace(Application.StartupPath & "\C_Autorizados.srf", "\\", "\")
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

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "frmA") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("frmA")
                oForm.Left = 400
                oForm.Visible = True
                oForm.PaneLevel = 1
                oForm.DataSources.DataTables.Add("MyDataTable")
            Else
                oForm = Me.SBO_Application.Forms.Item("frmA")
            End If
            cargarData()
            actualizarGrid()
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

    Private Sub visualizardata()
        Try
            Dim gridView As SAPbouiCOM.Grid
            gridView = oForm.Items.Item("Item_1").Specific
            Dim sql As String = "EXEC BUSCAR_INFO_DETALLE_RETENCION '" & UDT_UF.docEntry & "'"
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(sql)
            gridView.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            gridView.AutoResizeColumns()
            gridView.Columns.Item(0).Editable = False
            gridView.Columns.Item(1).Editable = False
            gridView.Columns.Item(2).Editable = False
            gridView.Columns.Item(3).Editable = False
            gridView.Columns.Item(4).Editable = False
            gridView.Columns.Item(5).Editable = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(gridView)
            gridView = Nothing
            GC.Collect()
        Catch ex As Exception
            Me.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
        End Try
    End Sub

    Private Sub cargarData()
        Try
            Dim oCombo1 As SAPbouiCOM.ComboBox
            Dim oCombo2 As SAPbouiCOM.ComboBox
            Dim orecord As SAPbobsCOM.Recordset
            orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCombo1 = oForm.Items.Item("Item_1").Specific
            oCombo2 = oForm.Items.Item("Item_4").Specific
            orecord.DoQuery("EXEC SP_CARGAR_DATOS '1','','',''")
            If orecord.RecordCount > 0 Then
                While orecord.EoF = False
                    oCombo1.ValidValues.Add(orecord.Fields.Item(0).Value, orecord.Fields.Item(1).Value)
                    orecord.MoveNext()
                End While
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
            orecord = Nothing
            GC.Collect()
            orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orecord.DoQuery("EXEC SP_CARGAR_DATOS '2','','',''")
            If orecord.RecordCount > 0 Then
                While orecord.EoF = False
                    oCombo2.ValidValues.Add(orecord.Fields.Item(0).Value, orecord.Fields.Item(1).Value)
                    orecord.MoveNext()
                End While
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
            orecord = Nothing
            GC.Collect()
        Catch ex As Exception
            SBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormTypeEx = "60004" And pVal.Before_Action = True And FormUID = "frmA" Then
                If pVal.ItemUID = "Item_5" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                    Dim orecord As SAPbobsCOM.Recordset
                    Dim oCombo1 As SAPbouiCOM.ComboBox
                    Dim oCombo2 As SAPbouiCOM.ComboBox
                    oCombo1 = oForm.Items.Item("Item_1").Specific
                    oCombo2 = oForm.Items.Item("Item_4").Specific
                    orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    orecord.DoQuery("EXEC SP_CARGAR_DATOS '3','" & oCombo1.Selected.Value & "','" & oCombo1.Selected.Description & "','" & oCombo2.Selected.Value & "'")
                    actualizarGrid()
                End If
            End If
        Catch ex As Exception
            Me.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try

    End Sub

    Private Sub actualizarGrid()
        Try
            Dim gridView As SAPbouiCOM.Grid
            gridView = oForm.Items.Item("Item_6").Specific
            'Dim sql As String = "EXEC BUSCAR_INFO_DETALLE_RETENCION '" & UDT_UF.docEntry & "'"
            oForm.DataSources.DataTables.Item(0).ExecuteQuery("SELECT * FROM [@COMPRO_AUTO] ORDER BY U_C_CODE ")
            gridView.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            gridView.AutoResizeColumns()
            gridView.Columns.Item(0).Visible = False
            gridView.Columns.Item(1).Visible = False
            gridView.Columns.Item(2).Editable = False
            gridView.Columns.Item(3).Editable = False
            gridView.Columns.Item(4).Editable = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(gridView)
            gridView = Nothing
            GC.Collect()
        Catch ex As Exception

        End Try
    End Sub

End Class
