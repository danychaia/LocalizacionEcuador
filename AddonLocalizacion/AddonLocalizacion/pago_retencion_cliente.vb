Public Class pago_retencion_cliente
    Private XmlForm As String = Replace(Application.StartupPath & "\pago_retenciones_clientes.srf", "\\", "\")
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter
    Private selected As Boolean = False
    Dim BaseTotal As Double = 0
    Dim RetencionTotal As Double = 0
    Public Sub New()
        MyBase.New()
        Try
            Me.SBO_Application = UDT_UF.SBOApplication
            Me.oCompany = UDT_UF.Company

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "pCliente") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("pCliente")
                oForm.Left = 400
                oForm.Visible = True
                oForm.PaneLevel = 1
                oForm.DataSources.DataTables.Add("MyDataTable")

            Else
                oForm = Me.SBO_Application.Forms.Item("pCliente")
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
            If pVal.FormUID = "pCliente" Then
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
                        If (pVal.ItemUID = "Item_23") Then
                            Try
                                Dim txtCuenta As SAPbouiCOM.EditText = oForm.Items.Item("Item_23").Specific
                                val = oDataTable.GetValue("FormatCode", 0)
                                txtCuenta.Value = val
                            Catch ex As Exception

                            End Try

                        End If
                        If (pVal.ItemUID = "Item_3") Then
                            Try
                                Dim txtFactura As SAPbouiCOM.EditText = oForm.Items.Item("Item_3").Specific
                                Dim txtBaseImponible As SAPbouiCOM.EditText = oForm.Items.Item("Item_14").Specific
                                Dim txtCliente As SAPbouiCOM.EditText = oForm.Items.Item("1").Specific
                                Dim txtImpuesto As SAPbouiCOM.EditText = oForm.Items.Item("Item_18").Specific
                                ' Dim txtNomEmp As SAPbouiCOM.EditText = oForm.Items.Item("1000008").Specific
                                val = oDataTable.GetValue("DocEntry", 0)
                                'UDT_UF.FilterCFL(oForm, "CFL_1", "DocEntry", val)
                                txtBaseImponible.Value = Double.Parse(obtenerBaseImponible(val)).ToString("N2")
                                txtCliente.Value = obtenerCliente(val)
                                txtImpuesto.Value = Double.Parse(obtenerImpuesto(val)).ToString("N2")
                                Try
                                    Dim txtRuc As SAPbouiCOM.EditText = oForm.Items.Item("1").Specific
                                    ' Dim txtNomEmp As SAPbouiCOM.EditText = oForm.Items.Item("1000008").Specific
                                    ' val = oDataTable.GetValue("CardCode", 0)
                                    Dim oBase As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_8").Specific
                                    Dim oretencion As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_13").Specific

                                    'UDT_UF.FilterCFL(oForm, "CFL_1", "DocEntry", val)
                                    Dim sql As String = "exec INF_PARTNER_OPE 2,'" & txtRuc.Value & "','','',''"
                                    Try
                                        Dim orecord As SAPbobsCOM.Recordset
                                        orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        orecord.DoQuery(sql)
                                        If orecord.RecordCount > 0 Then
                                            While orecord.EoF = False
                                                oBase.ValidValues.Add(orecord.Fields.Item(3).Value, "%")
                                                oretencion.ValidValues.Add(orecord.Fields.Item(4).Value, "%")
                                                orecord.MoveNext()
                                            End While
                                            oBase.ValidValues.Add("0", "%")
                                            oretencion.ValidValues.Add("0", "%")
                                        End If
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                                        orecord = Nothing
                                        GC.Collect()
                                    Catch ex As Exception

                                    End Try
                                Catch ex As Exception

                                End Try
                                txtFactura.Value = val
                            Catch ex As Exception

                            End Try

                        End If
                    End If
                End If               
            End If

            If pVal.Before_Action = True And pVal.FormUID = "pCliente" And pVal.ItemUID = "Item_22" Then
                Dim oBase As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_8").Specific
                Dim oRetencion As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_13").Specific
                Dim txtBaseImponible As SAPbouiCOM.EditText = oForm.Items.Item("Item_14").Specific
                Dim Impuesto As SAPbouiCOM.EditText = oForm.Items.Item("Item_18").Specific
                Dim txtTotalB As SAPbouiCOM.EditText = oForm.Items.Item("Item_6").Specific
                Dim txtTotalR As SAPbouiCOM.EditText = oForm.Items.Item("Item_10").Specific
                Dim txtCliente As SAPbouiCOM.EditText = oForm.Items.Item("1").Specific
                Dim txtDocumento As SAPbouiCOM.EditText = oForm.Items.Item("Item_3").Specific
                Dim orecord As SAPbobsCOM.Recordset
                orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If validar(1) = False Then
                    BubbleEvent = False
                    Return
                End If
                Dim sql As String = "INF_PAGO_RETENCION '1','" & txtCliente.Value & "','" & txtDocumento.Value & "'," & oBase.Value.Trim & "," & oRetencion.Value & "," & (Double.Parse(oBase.Value) / 100) * Double.Parse(txtBaseImponible.Value) & "," & (Double.Parse(oRetencion.Value) / 100) * Double.Parse(Impuesto.Value)
                orecord.DoQuery(sql)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                orecord = Nothing
                GC.Collect()
                cargarRetenciones(txtCliente.Value, txtDocumento.Value)
                txtTotalB.Value = Double.Parse(BaseTotal).ToString("N2")
                txtTotalR.Value = Double.Parse(RetencionTotal).ToString("N2")
                BaseTotal = 0
                RetencionTotal = 0
                BubbleEvent = False
                Return
            End If

            If pVal.Before_Action = True And pVal.FormUID = "pCliente" And pVal.ItemUID = "Item_20" Then
                Dim txtCliente As SAPbouiCOM.EditText = oForm.Items.Item("1").Specific
                Dim txtDocumento As SAPbouiCOM.EditText = oForm.Items.Item("Item_3").Specific
                Dim txtTotalB As SAPbouiCOM.EditText = oForm.Items.Item("Item_6").Specific
                Dim txtTotalR As SAPbouiCOM.EditText = oForm.Items.Item("Item_10").Specific
                If validar(2) = False Then
                    BubbleEvent = False
                    Return
                End If
                cargarRetenciones(txtCliente.Value, txtDocumento.Value)
                txtTotalB.Value = Double.Parse(BaseTotal).ToString("N2")
                txtTotalR.Value = Double.Parse(RetencionTotal).ToString("N2")
                BaseTotal = 0
                RetencionTotal = 0
                BubbleEvent = False
                Return
            End If
            If pVal.Before_Action = True And pVal.FormUID = "pCliente" And pVal.ItemUID = "Item_21" Then
                Try
                    Dim txtDocumento As SAPbouiCOM.EditText = oForm.Items.Item("Item_3").Specific
                    Dim txtCliente As SAPbouiCOM.EditText = oForm.Items.Item("1").Specific
                    Dim txtTotalB As SAPbouiCOM.EditText = oForm.Items.Item("Item_6").Specific
                    Dim txtTotalR As SAPbouiCOM.EditText = oForm.Items.Item("Item_10").Specific
                    Dim txtcuenta As SAPbouiCOM.EditText = oForm.Items.Item("Item_23").Specific
                   
                    If validar(3) = False Then
                        BubbleEvent = False
                        Return
                    End If
                    Dim InPay As SAPbobsCOM.Payments
                    'Dim oDownPay As SAPbobsCOM.Documents
                    'oDownPay = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)                  
                    InPay = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

                    ' oDownPay.GetByKey(Convert.ToInt32(sNewObjCode))

                    InPay.CardCode = txtCliente.Value.Trim

                    InPay.Invoices.DocEntry = txtDocumento.Value.Trim
                    InPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                    InPay.Remarks = "Pago de Retencion Localizacion"
                    If Double.Parse(txtTotalB.Value) > 0 Then
                        InPay.CreditCards.CreditCard = 1  ' Mastercard = 1 , VISA = 2
                        InPay.CreditCards.CardValidUntil = CDate("01/12/2020")
                        InPay.CreditCards.CreditCardNumber = "1220" ' Just need 4 last digits
                        InPay.CreditCards.CreditSum = txtTotalB.Value   ' Total Amount of the Invoice
                        InPay.CreditCards.VoucherNum = "1234567" ' Need to give the Credit Card confirmation number.
                        InPay.CreditCards.PaymentMethodCode = 1
                    End If
                    If Double.Parse(txtTotalR.Value) > 0 Then
                        InPay.CreditCards.Add()
                        InPay.CreditCards.CreditCard = 2  ' Mastercard = 1 , VISA = 2
                        InPay.CreditCards.CardValidUntil = CDate("01/12/2020")
                        InPay.CreditCards.CreditCardNumber = "1220" ' Just need 4 last digits
                        InPay.CreditCards.CreditSum = Double.Parse(txtTotalR.Value)  ' Total Amount of the Invoice
                        InPay.CreditCards.VoucherNum = "1234567" ' Need to give the Credit Card confirmation number.
                        InPay.CreditCards.PaymentMethodCode = 2
                    End If
                    
                    If InPay.Add() <> 0 Then
                        SBOApplication.SetStatusBarMessage(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        MsgBox(oCompany.GetLastErrorDescription())
                    Else
                        SBOApplication.SetStatusBarMessage("Pago de retencion correcto!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        oForm.Close()
                    End If
                    BubbleEvent = False
                    Return
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
                BubbleEvent = False
                Return
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub
    Private Sub visualizardata(p1 As String, p2 As String)
        Try

            Dim gridView As SAPbouiCOM.Grid
            gridView = oForm.Items.Item("Item_6").Specific
            Dim sql As String = "EXEC BUSCAR_INFO_RETENCION '" & p1 & "','" & p2 & "'"
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(sql)
            gridView.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            gridView.AutoResizeColumns()
            gridView.Columns.Item(0).Visible = False
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

    Private Function obtenerBaseImponible(val As String) As String
        Dim baseImponible As String = ""
        Try
            Dim orecord As SAPbobsCOM.Recordset
            orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sql As String = "SELECT SUM(A.LineTotal) FROM INV1 A WHERE A.DocEntry = '" & val & "'"
            orecord.DoQuery(sql)
            If orecord.RecordCount > 0 Then
                Return orecord.Fields.Item(0).Value
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
            orecord = Nothing
            GC.Collect()
        Catch ex As Exception

        End Try
        Return baseImponible
    End Function

    Private Function obtenerCliente(val As String) As String
        Dim baseImponible As String = ""
        Try
            Dim orecord As SAPbobsCOM.Recordset
            orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sql As String = "SELECT A.CardCode FROM OINV A WHERE A.DocEntry = '" & val & "'"
            orecord.DoQuery(sql)
            If orecord.RecordCount > 0 Then
                Return orecord.Fields.Item(0).Value
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
            orecord = Nothing
            GC.Collect()
        Catch ex As Exception

        End Try
        Return baseImponible
    End Function

    Private Function obtenerImpuesto(val As String) As String
        Dim impuesto As String = ""
        Try
            Dim orecord As SAPbobsCOM.Recordset
            orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sql As String = "SELECT SUM(A.VatSum) FROM INV1 A WHERE A.DocEntry = '" & val & "'"
            orecord.DoQuery(sql)
            If orecord.RecordCount > 0 Then
                Return orecord.Fields.Item(0).Value
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
            orecord = Nothing
            GC.Collect()
        Catch ex As Exception

        End Try
        Return impuesto
    End Function

    Private Sub cargarRetenciones(p1 As String, p2 As String)
        Try
            Dim gridView As SAPbouiCOM.Grid
            gridView = oForm.Items.Item("Item_19").Specific
            Dim sql As String = "INF_PAGO_RETENCION '2','" & p1 & "','" & p2 & "',0,0,0,0"
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(sql)
            gridView.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            gridView.AutoResizeColumns()
            gridView.Columns.Item(0).Visible = False
            gridView.Columns.Item(1).Editable = False
            gridView.Columns.Item(2).Editable = False
            gridView.Columns.Item(3).Editable = False
            gridView.Columns.Item(4).Editable = False
            gridView.Columns.Item(5).Editable = True
            gridView.Columns.Item(6).Editable = True
            For index = 0 To gridView.Rows.Count - 1
                BaseTotal += Double.Parse(gridView.DataTable.GetValue(gridView.DataTable.Columns.Item(5).Name, index))
                RetencionTotal += Double.Parse(gridView.DataTable.GetValue(gridView.DataTable.Columns.Item(6).Name, index))
            Next
            System.Runtime.InteropServices.Marshal.ReleaseComObject(gridView)
            gridView = Nothing
            GC.Collect()

        Catch ex As Exception
            Me.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
        End Try
    End Sub

    Private Function validar(tipo As Integer) As Boolean

        Try
            Dim oBase As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_8").Specific
            Dim oRetencion As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_13").Specific
            Dim txtBaseImponible As SAPbouiCOM.EditText = oForm.Items.Item("Item_14").Specific
            Dim Impuesto As SAPbouiCOM.EditText = oForm.Items.Item("Item_18").Specific
            Dim txtTotalB As SAPbouiCOM.EditText = oForm.Items.Item("Item_6").Specific
            Dim txtTotalR As SAPbouiCOM.EditText = oForm.Items.Item("Item_10").Specific
            Dim txtCliente As SAPbouiCOM.EditText = oForm.Items.Item("1").Specific
            Dim txtDocumento As SAPbouiCOM.EditText = oForm.Items.Item("Item_3").Specific
            Dim txtCuenta As SAPbouiCOM.EditText = oForm.Items.Item("Item_23").Specific
            If tipo = 1 Then
                If txtDocumento.Value = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de seleccionar un Documento", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                If oBase.Value.Trim = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de seleccionar una Base", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                If oRetencion.Value.Trim = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de seleccionar una Retención", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
            End If
            If tipo = 2 Then
                If txtDocumento.Value = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de seleccionar un Documento", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
            End If

            If tipo = 3 Then
                If txtDocumento.Value = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de seleccionar un Documento", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                If txtTotalB.Value.Trim = "0.00" And txtTotalR.Value = "0.00" Then
                    SBO_Application.SetStatusBarMessage("El monto a pagar debe se mayor a 0.00", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                If txtTotalB.Value.Trim = "" And txtTotalR.Value = "" Then
                    SBO_Application.SetStatusBarMessage("El existe monto a pagar", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                ' If txtCuenta.Value = "" Then
                'SBO_Application.SetStatusBarMessage("Debe de seleccionar una Cuenta", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                '  Return False
                ' End If
            End If

        Catch ex As Exception

        End Try

        Return True
    End Function
End Class
