Public Class retencion_numeros
    Private XmlForm As String = Replace(Application.StartupPath & "\series_detalle.srf", "\\", "\")
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter
    Public code As String = ""
    Public Sub New()
        MyBase.New()
        Try
            Me.SBO_Application = UDT_UF.SBOApplication
            Me.oCompany = UDT_UF.Company
            UDT_UF.code = ""
            If UDT_UF.ActivateFormIsOpen(SBO_Application, "fSdt") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("fSdt")
                oForm.Left = 400
                oForm.Visible = True
                oForm.PaneLevel = 1

                'Dim oEstable As SAPbouiCOM.EditText
                'Dim oPunto As SAPbouiCOM.EditText
                'Dim oDeI As SAPbouiCOM.EditText
                'Dim oHastaI As SAPbouiCOM.EditText
                'Dim oCombo As SAPbouiCOM.ComboBox
                'Dim cmdenviar As SAPbouiCOM.Button
                'esto es para poder hacer que los textos tengan formato de fecha
                oForm.DataSources.DataTables.Add("MyDataTable")
                'oForm.DataSources.UserDataSources.Add("Date", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                'oForm.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                'oForm.DataSources.UserDataSources.Add("De", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                'oForm.DataSources.UserDataSources.Add("Hasta", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                'oEstable = oForm.Items.Item("Item_2").Specific
                'oPunto = oForm.Items.Item("Item_4").Specific
                'oDeI = oForm.Items.Item("Item_9").Specific
                'oHastaI = oForm.Items.Item("Item_11").Specific
                'oEstable.DataBind.SetBound(True, "", "Date")
                'oPunto.DataBind.SetBound(True, "", "Date2")
                'oCombo = oForm.Items.Item("Item_8").Specific
                'oCombo.ValidValues.Add("01", "Electrónico")
                'oCombo.ValidValues.Add("02", "Impreso")

            Else
                oForm = Me.SBO_Application.Forms.Item("fSdt")
            End If
            carcarSeries()
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
            If pVal.FormTypeEx = "60004" And pVal.Before_Action = True And pVal.FormUID = "fSdt" Then
                If pVal.ItemUID = "Item_5" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                    Dim obutton As SAPbouiCOM.Button
                    obutton = oForm.Items.Item("Item_5").Specific
                    Dim oDe As SAPbouiCOM.EditText
                    Dim oHasta As SAPbouiCOM.EditText
                    Dim oCombo As SAPbouiCOM.ComboBox
                    Dim oDeI As SAPbouiCOM.EditText
                    Dim oHastaI As SAPbouiCOM.EditText
                    oDe = oForm.Items.Item("Item_2").Specific
                    oHasta = oForm.Items.Item("Item_4").Specific
                    oCombo = oForm.Items.Item("Item_8").Specific
                    oDeI = oForm.Items.Item("Item_9").Specific
                    oHastaI = oForm.Items.Item("Item_11").Specific
                    If obutton.Caption.Equals("Agregar") Then

                        If oDe.Value = "" Or oHasta.Value = "" Or oCombo.Value = "" Then
                            SBOApplication.SetStatusBarMessage("Debe de ingresar una Serie de Establecimiento, un punto de emisión y un Tipo de Documento", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        Else
                            If oCombo.Selected.Description = "Impreso" Then
                                If oDeI.Value = "" Or oHastaI.Value = "" Then
                                    SBO_Application.SetStatusBarMessage("Para Facturas Impresas debe de Ingresar un rango", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    BubbleEvent = False
                                    Return
                                End If
                            End If
                            Dim orecord As SAPbobsCOM.Recordset
                            orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim sql As String = "EXEC SERIES_PTO_ESTABLE '1','" & oDe.Value.PadLeft(3, "0") & "','" & oHasta.Value.PadLeft(3, "0") & "','','" & oCombo.Selected.Description & "','" & oDeI.Value & "','" & oHastaI.Value & "'"
                            orecord.DoQuery(sql)
                            carcarSeries()
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                            orecord = Nothing
                            GC.Collect()
                        End If
                    Else
                        If obutton.Caption.Equals("Eliminar") Then
                            If code <> "" Then
                                Dim orecord As SAPbobsCOM.Recordset
                                orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim sql As String = "EXEC SERIES_PTO_ESTABLE '2','" & oDe.Value & "','" & oHasta.Value & "','" & code & "','','',''"
                                orecord.DoQuery(sql)
                                carcarSeries()
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                                orecord = Nothing
                                GC.Collect()
                            Else
                                SBO_Application.SetStatusBarMessage("Debe de seleccionar una fila", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            End If
                        End If
                    End If

                End If
                If pVal.ItemUID = "Item_0" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                    Try
                        Dim gridView As SAPbouiCOM.Grid
                        gridView = oForm.Items.Item("Item_0").Specific
                        If pVal.Row <> -1 Then
                            code = gridView.DataTable.GetValue("Code", pVal.Row).ToString
                            Dim estable = gridView.DataTable.GetValue(1, pVal.Row).ToString
                            Dim punto = gridView.DataTable.GetValue(2, pVal.Row).ToString
                            Dim correlativo = Integer.Parse(gridView.DataTable.GetValue(4, pVal.Row).ToString)
                            'correlativo += 1
                            'Dim obutton As SAPbouiCOM.Button
                            ' obutton = oForm.Items.Item("Item_5").Specific
                            'obutton.Caption = "Eliminar"
                            Dim oCompuesto = estable & punto & correlativo.ToString.PadLeft(7, "0")
                            UDT_UF.code = oCompuesto & "-" & code
                            oForm.Close()
                            'Dim detalle As New retencion_info_detalle
                            BubbleEvent = False
                            Return
                        Else
                            BubbleEvent = False
                            Return
                        End If

                    Catch ex As Exception

                    End Try
                  
                End If
                If (pVal.ItemUID = "Item_2" Or pVal.ItemUID = "Item_4") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Before_Action = True Then
                    Dim obutton As SAPbouiCOM.Button
                    obutton = oForm.Items.Item("Item_5").Specific
                    obutton.Caption = "Agregar"
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message)
        End Try

    End Sub
    Private Sub carcarSeries()
        Try
            Dim gridView As SAPbouiCOM.Grid
            gridView = oForm.Items.Item("Item_0").Specific
            gridView.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            Dim sql As String = "EXEC SERIES_PTO_ESTABLE '3','','','','','','','','',''"
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
