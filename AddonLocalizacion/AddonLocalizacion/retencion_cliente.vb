Public Class retencion_cliente
    Private XmlForm As String = Replace(Application.StartupPath & "\retencion_clientes.srf", "\\", "\")
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

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "rCliente") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("rCliente")
                oForm.Left = 400
                oForm.Visible = True
                oForm.PaneLevel = 1
                oForm.DataSources.DataTables.Add("MyDataTable")

            Else
                oForm = Me.SBO_Application.Forms.Item("rCliente")
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
        Try
            If pVal.FormUID = "rCliente" Then                
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBO_Application.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                    If oCFLEvento.BeforeAction = False Then
                        Dim oDataTable As SAPbouiCOM.DataTable
                        oDataTable = oCFLEvento.SelectedObjects
                        Dim val As String
                        If (pVal.ItemUID = "Item_1") Then
                            Try
                                Dim txtRuc As SAPbouiCOM.EditText = oForm.Items.Item("Item_1").Specific
                                Dim txtNombre As SAPbouiCOM.EditText = oForm.Items.Item("Item_3").Specific
                                ' Dim txtNomEmp As SAPbouiCOM.EditText = oForm.Items.Item("1000008").Specific
                                val = oDataTable.GetValue("CardCode", 0)
                                txtNombre.Value = obtenerNombre(val)
                                txtRuc.Value = val
                            Catch ex As Exception

                            End Try

                        End If
                    End If

                End If
            End If
            If pVal.FormUID = "rCliente" And pVal.Before_Action = True And pVal.ItemUID = "Item_11" Then
                Dim txtCliente As SAPbouiCOM.EditText = oForm.Items.Item("Item_1").Specific
                Dim txtRazon As SAPbouiCOM.EditText = oForm.Items.Item("Item_3").Specific
                Dim txtBase As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_4").Specific
                Dim txtRetencion As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_10").Specific

                If txtCliente.Value = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de Seleccionar un cliente", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False
                    Return
                End If
                If txtBase.Value.Trim = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de Seleccionar una base.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False
                    Return
                End If
                If txtRetencion.Value = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de Seleccionar una retención", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False
                    Return
                End If
                Dim orecord As SAPbobsCOM.Recordset
                orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim sql As String = "exec INF_PARTNER_OPE 1,'" & txtCliente.Value & "','" & txtRazon.Value & "','" & txtBase.Value.Trim & "','" & txtRetencion.Value.Trim & "'"
                orecord.DoQuery(sql)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                orecord = Nothing
                GC.Collect()
                carcarSeries(txtCliente.Value)
            End If

            If pVal.FormUID = "rCliente" And pVal.Before_Action = True And pVal.ItemUID = "Item_12" Then
                Dim txtCliente As SAPbouiCOM.EditText = oForm.Items.Item("Item_1").Specific
                If txtCliente.Value = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de Seleccionar un cliente", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False
                    Return
                End If

                carcarSeries(txtCliente.Value)
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub

    Private Function obtenerNombre(val As String) As String
        Dim nombre As String = ""
        Try
            Dim orecord As SAPbobsCOM.Recordset
            orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sql As String = "SELECT A.CardName FROM OCRD A where A.CardCode = '" & val & "'"
            orecord.DoQuery(sql)
            If orecord.RecordCount > 0 Then
                Return orecord.Fields.Item(0).Value
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
            orecord = Nothing
            GC.Collect()
        Catch ex As Exception

        End Try
        Return nombre
    End Function

    Private Sub carcarSeries(cliente As String)
        Try
            Dim gridView As SAPbouiCOM.Grid
            gridView = oForm.Items.Item("Item_9").Specific
            gridView.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            Dim sql As String = "exec INF_PARTNER_OPE 2,'" & cliente & "','','',''"
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(sql)
            gridView.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            gridView.AutoResizeColumns()
            gridView.Columns.Item(0).Visible = False
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
