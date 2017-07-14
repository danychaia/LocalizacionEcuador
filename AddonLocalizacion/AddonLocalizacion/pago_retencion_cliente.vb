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
                Dim txtFechaReten As SAPbouiCOM.EditText = oForm.Items.Item("Item_29").Specific
                Dim txtCaducidad As SAPbouiCOM.EditText = oForm.Items.Item("Item_31").Specific
                oForm.DataSources.UserDataSources.Add("Date", SAPbouiCOM.BoDataType.dt_DATE)
                oForm.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_DATE)
                txtFechaReten.DataBind.SetBound(True, "", "Date")
                txtCaducidad.DataBind.SetBound(True, "", "Date2")

            Else
                oForm = Me.SBO_Application.Forms.Item("pCliente")
            End If
            cargarcombo()
            cargarSeriesSAP()
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
                                Dim oRecordB As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Try
                                    'UDT_UF.FilterCFL(oForm, "CFL_1", "DocEntry", val)
                                    Dim txtAutorizacion As SAPbouiCOM.EditText = oForm.Items.Item("Item_24").Specific
                                    Dim txtNumeroReten As SAPbouiCOM.EditText = oForm.Items.Item("Item_27").Specific
                                    Dim txtFechaReten As SAPbouiCOM.EditText = oForm.Items.Item("Item_29").Specific
                                    Dim txtCaducidad As SAPbouiCOM.EditText = oForm.Items.Item("Item_31").Specific
                                    val = oDataTable.GetValue("DocEntry", 0)
                                    txtBaseImponible.Value = Double.Parse(obtenerBaseImponible(val)).ToString("N2")
                                    txtCliente.Value = obtenerCliente(val)
                                    txtImpuesto.Value = Double.Parse(obtenerImpuesto(val)).ToString("N2")
                                    oRecordB.DoQuery("EXEC INF_PAGO_RETENCION '3','" & txtCliente.Value & "','" & val & "',0,0,0,0,'','',''")
                                    If oRecordB.Fields.Item(0).Value = "YA EXISTE" Then
                                        cargarRetenciones(txtCliente.Value.Trim, val)
                                        SBOApplication.SetStatusBarMessage("A este documento ya se le ha aplicado pago de retención", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    ElseIf oRecordB.Fields.Item(0).Value = "NO EXISTE" Then
                                        cargarRetenciones(txtCliente.Value.Trim, val)
                                    End If
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordB)
                                    oRecordB = Nothing
                                    GC.Collect()
                                    Dim txtTotalB As SAPbouiCOM.EditText = oForm.Items.Item("Item_6").Specific
                                    Dim txtTotalR As SAPbouiCOM.EditText = oForm.Items.Item("Item_10").Specific
                                    Dim total As SAPbobsCOM.Recordset
                                    total = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    total.DoQuery("INF_PAGO_RETENCION '4','" & txtCliente.Value & "','" & val & "',0,0,0,0,'','',''")
                                    txtTotalB.Value = Double.Parse(total.Fields.Item(0).Value)
                                    txtTotalR.Value = Double.Parse(total.Fields.Item(1).Value)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(total)
                                    total = Nothing
                                    GC.Collect()
                                    total = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    total.DoQuery("EXEC INFORMACION_PAGO '2','','','','','" & txtCliente.Value & "','" & val & "'")
                                    If total.RecordCount > 0 Then
                                        txtAutorizacion.Value = total.Fields.Item(2).Value
                                        txtNumeroReten.Value = total.Fields.Item(3).Value
                                        txtFechaReten.Value = Date.Parse(total.Fields.Item(4).Value).ToString("yyyyMMdd")
                                        txtCaducidad.Value = Date.Parse(total.Fields.Item(5).Value).ToString("yyyyMMdd")
                                    End If
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(total)
                                    total = Nothing
                                    GC.Collect()
                                Catch ex As Exception

                                End Try

                               
                                Try
                                    Dim txtRuc As SAPbouiCOM.EditText = oForm.Items.Item("1").Specific
                                    ' Dim txtNomEmp As SAPbouiCOM.EditText = oForm.Items.Item("1000008").Specific
                                    ' val = oDataTable.GetValue("CardCode", 0)
                                    Dim oBase As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_8").Specific
                                    Dim oretencion As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_13").Specific

                                    'UDT_UF.FilterCFL(oForm, "CFL_1", "DocEntry", val)
                                    Dim sql As String = "exec INF_PARTNER_OPE 2,'" & txtRuc.Value & "','','','','',''"
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
                        If pVal.ItemUID = "Item_32" Then
                            Try
                                Dim txtCuenta As SAPbouiCOM.EditText = oForm.Items.Item("Item_32").Specific

                                ' Dim txtNomEmp As SAPbouiCOM.EditText = oForm.Items.Item("1000008").Specific
                                val = oDataTable.GetValue("CreditCard", 0)

                                txtCuenta.Value = val
                            Catch ex As Exception

                            End Try
                        End If
                        If pVal.ItemUID = "Item_33" Then
                            Try
                                Dim txtCuenta As SAPbouiCOM.EditText = oForm.Items.Item("Item_33").Specific

                                ' Dim txtNomEmp As SAPbouiCOM.EditText = oForm.Items.Item("1000008").Specific
                                val = oDataTable.GetValue("CreditCard", 0)

                                txtCuenta.Value = val
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
                Dim txtCuentaB As SAPbouiCOM.EditText = oForm.Items.Item("Item_32").Specific
                Dim txtCuentaR As SAPbouiCOM.EditText = oForm.Items.Item("Item_33").Specific

                Dim orecord As SAPbobsCOM.Recordset
                orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If validar(1) = False Then
                    BubbleEvent = False
                    Return
                End If
              
                Dim sql As String = "INF_PAGO_RETENCION '1','" & txtCliente.Value & "','" & txtDocumento.Value & "'," & oBase.Value.Trim & "," & oRetencion.Value & "," & (Double.Parse(oBase.Value) / 100) * Double.Parse(txtBaseImponible.Value) & "," & (IIf(oRetencion.Value.Trim = "", 0, Double.Parse(oRetencion.Value)) / 100) * Double.Parse(Impuesto.Value) & ",'" & oBase.Selected.Description & "','" & oRetencion.Selected.Description & "','C'"
                orecord.DoQuery(sql)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                orecord = Nothing
                GC.Collect()
                cargarRetenciones(txtCliente.Value, txtDocumento.Value)

                sql = "INF_PAGO_RETENCION '4','" & txtCliente.Value & "','" & txtDocumento.Value & "'," & oBase.Value.Trim & "," & oRetencion.Value & "," & (Double.Parse(oBase.Value) / 100) * Double.Parse(txtBaseImponible.Value) & "," & (IIf(oRetencion.Value.Trim = "", 0, Double.Parse(oRetencion.Value)) / 100) * Double.Parse(Impuesto.Value) & ",'" & txtCuentaB.Value & "','" & txtCuentaR.Value & "','C'"
                orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                orecord.DoQuery(sql)
                             
                txtTotalB.Value = Double.Parse(orecord.Fields.Item(0).Value)
                txtTotalR.Value = Double.Parse(orecord.Fields.Item(1).Value)
                BaseTotal = 0
                RetencionTotal = 0
                txtCuentaB.Value = oBase.Selected.Description.Trim
                txtCuentaR.Value = oRetencion.Selected.Description.Trim
                System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                orecord = Nothing
                GC.Collect()
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
                    Dim txtCuentaB As SAPbouiCOM.EditText = oForm.Items.Item("Item_32").Specific
                    Dim txtCuentaR As SAPbouiCOM.EditText = oForm.Items.Item("Item_33").Specific
                    Dim oSeries As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_40").Specific
                    Dim txtAutorizacion As SAPbouiCOM.EditText = oForm.Items.Item("Item_24").Specific
                    Dim txtNumeroReten As SAPbouiCOM.EditText = oForm.Items.Item("Item_27").Specific
                    Dim txtFechaReten As SAPbouiCOM.EditText = oForm.Items.Item("Item_29").Specific
                    Dim txtCaducidad As SAPbouiCOM.EditText = oForm.Items.Item("Item_31").Specific
                    Dim oRecordv As SAPbobsCOM.Recordset
                    Dim orecordPago As SAPbobsCOM.Recordset


                    oRecordv = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordv.DoQuery("EXEC INF_PAGO_CLIENTE '1','" & txtDocumento.Value & "','" & txtCliente.Value & "'")
                    If oRecordv.Fields.Item(0).Value = "NO" Then
                        SBO_Application.SetStatusBarMessage("el Pago ya fue realizado", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False
                        Return
                    End If

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordv)
                    oRecordv = Nothing
                    GC.Collect()
                    If validar(3) = False Then
                        BubbleEvent = False
                        Return
                    End If
                    If UDT_UF.infoPago Is Nothing Then
                        Dim info As New Inf_pago_retencion
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
                    InPay.Remarks = UDT_UF.infoPago.remarks
                    InPay.Series = oSeries.Value
                    orecordPago = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim pago As Double = 0
                    Dim valor As String
                    Dim sqls = "INF_PAGO_RETENCION '5','" & txtCliente.Value & "','" & txtDocumento.Value & "',0,0,0,0,'','',''"
                    orecordPago.DoQuery(sqls)
                    While orecordPago.EoF = False
                        If Double.Parse(orecordPago.Fields.Item("U_T_BASE").Value) > 0 Then
                            InPay.CreditCards.CreditCard = orecordPago.Fields.Item("U_CUENTAB").Value    ' Mastercard = 1 , VISA = 2
                            InPay.CreditCards.CardValidUntil = UDT_UF.infoPago.validDate
                            InPay.CreditCards.CreditCardNumber = UDT_UF.infoPago.creditNum  ' Just need 4 last digits
                            valor = orecordPago.Fields.Item("U_T_BASE").Value
                            pago = pago + Double.Parse(orecordPago.Fields.Item("U_T_BASE").Value)
                            InPay.CreditCards.CreditSum = Double.Parse(orecordPago.Fields.Item("U_T_BASE").Value)   ' Total Amount of the Invoice
                            InPay.CreditCards.VoucherNum = UDT_UF.infoPago.voucher  ' Need to give the Credit Card confirmation number.
                            Dim orecord As SAPbobsCOM.Recordset
                            orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orecord.DoQuery("EXEC SP_CUENTA_RETENCIONES 'Pago','" & orecordPago.Fields.Item("U_CUENTAB").Value & "'")
                            InPay.CreditCards.PaymentMethodCode = orecord.Fields.Item(0).Value
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                            orecord = Nothing
                            GC.Collect()
                            InPay.CreditCards.Add()
                        End If
                        orecordPago.MoveNext()
                    End While
                    InPay.Invoices.SumApplied = pago
                    If InPay.Add() <> 0 Then
                        SBOApplication.SetStatusBarMessage(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Else
                        SBOApplication.SetStatusBarMessage("Pago de retencion correcto!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        UDT_UF.infoPago = Nothing
                        oRecordv = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordv.DoQuery("EXEC INF_PAGO_CLIENTE '2','" & txtDocumento.Value & "','" & txtCliente.Value & "'")
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordv)
                        oRecordv = Nothing
                        GC.Collect()
                        oRecordv = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordv.DoQuery(" EXEC INFORMACION_PAGO '1','" & txtAutorizacion.Value & "','" & txtNumeroReten.Value & "','" & txtFechaReten.Value & "','" & txtCaducidad.Value & "','" & txtCliente.Value & "','" & txtDocumento.Value & "'")
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordv)
                        oRecordv = Nothing
                        GC.Collect()
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

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.BeforeAction = True Then
                Me.oForm = Nothing
                Me.oCompany = Nothing
                Me.SBO_Application = Nothing
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
            Dim sql As String = "INF_PAGO_RETENCION '2','" & p1 & "','" & p2 & "',0,0,0,0,'','',''"
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
            Dim txtAutorizacion As SAPbouiCOM.EditText = oForm.Items.Item("Item_24").Specific
            Dim txtNumeroReten As SAPbouiCOM.EditText = oForm.Items.Item("Item_27").Specific
            Dim txtFechaReten As SAPbouiCOM.EditText = oForm.Items.Item("Item_29").Specific
            Dim txtCaducidad As SAPbouiCOM.EditText = oForm.Items.Item("Item_31").Specific
            Dim txtCuentaB As SAPbouiCOM.EditText = oForm.Items.Item("Item_32").Specific
            Dim txtCuentaR As SAPbouiCOM.EditText = oForm.Items.Item("Item_33").Specific
            Dim oSeries As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_40").Specific
            If tipo = 1 Then
                If txtDocumento.Value = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de seleccionar un Documento", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                If oBase.Value.Trim = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de seleccionar una Base", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If

                If oSeries.Value.Trim = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de seleccionar una Serie para continuar", SAPbouiCOM.BoMessageTime.bmt_Short, True)
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

                If txtNumeroReten.Value.Trim = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de seleccionar un numero de retención", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                If txtAutorizacion.Value.Trim = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de ingresar una autorización", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                If txtFechaReten.Value.Trim = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de ingresar una fecha de retencion", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                If txtCaducidad.Value.Trim = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de ingresar una fecha de Caducidad", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                'If txtCuentaB.Value.Trim = "" And Double.Parse(txtTotalB.Value) > 0 Then
                'SBO_Application.SetStatusBarMessage("Debe seleccionar una cuenta como base", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                ' Return False
                'End If
                'If txtCuentaR.Value.Trim = "" And Double.Parse(txtTotalR.Value) > 0 Then
                'SBO_Application.SetStatusBarMessage("Debe seleccionar una cuenta como Retencion", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                'Return False
                ' End If
                ' If txtCuenta.Value = "" Then
                'SBO_Application.SetStatusBarMessage("Debe de seleccionar una Cuenta", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                '  Return False
                ' End If
            End If


        Catch ex As Exception

        End Try

        Return True
    End Function

    Private Sub cargarcombo()
        Try
            Dim oComboFuente As SAPbouiCOM.ComboBox
            Dim oComboIVA As SAPbouiCOM.ComboBox
            Dim oRecord As SAPbobsCOM.Recordset

            oComboFuente = oForm.Items.Item("Item_8").Specific
            oComboIVA = oForm.Items.Item("Item_13").Specific

            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery(" exec SP_CUENTA_RETENCIONES 'Fuente',''")
            If oRecord.RecordCount > 0 Then
                oComboFuente.ValidValues.Add("0", "%")
                While oRecord.EoF = False
                    oComboFuente.ValidValues.Add(oRecord.Fields.Item(0).Value, oRecord.Fields.Item(1).Value)
                    oRecord.MoveNext()
                End While
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()

            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery(" exec SP_CUENTA_RETENCIONES 'IVA',''")
            oComboIVA.ValidValues.Add("0", "%")
            If oRecord.RecordCount > 0 Then
                While oRecord.EoF = False
                    oComboIVA.ValidValues.Add(oRecord.Fields.Item(0).Value, oRecord.Fields.Item(1).Value)
                    oRecord.MoveNext()
                End While
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub cargarSeriesSAP()
        Try
            Dim oSeries As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_40").Specific
            oSeries.ValidValues.LoadSeries("24", SAPbouiCOM.BoSeriesMode.sf_View)
            'Dim oSeriesCombo As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_40").Specific
            'Dim oRecord As SAPbobsCOM.Recordset
            'oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ' oRecord.DoQuery("SP_SERIES_PAGO_VENTA")
            'If oRecord.RecordCount > 0 Then
            'While oRecord.EoF = False
            'oSeriesCombo.ValidValues.Add(oRecord.Fields.Item(0).Value, oRecord.Fields.Item(1).Value)
            ' oRecord.MoveNext()
            ' End While
            'End If
        Catch ex As Exception
        End Try
    End Sub

End Class
