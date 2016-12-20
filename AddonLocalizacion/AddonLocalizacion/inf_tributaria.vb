Public Class inf_tributaria
    Private XmlForm As String = Replace(Application.StartupPath & "\inf_tributaria.srf", "\\", "\")
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
            Dim ruc As SAPbouiCOM.EditText
            Dim estable As SAPbouiCOM.EditText
            Dim ptoEmisor As SAPbouiCOM.EditText
            Dim ofolder As SAPbouiCOM.Folder
            If UDT_UF.ActivateFormIsOpen(SBO_Application, "frm_inf") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("frm_inf")
                oForm.Visible = True
                'oForm.PaneLevel = 1
                'oForm.Left = 20
                ruc = oForm.Items.Item("txtruc").Specific
                ruc.Value = "0".PadRight(13, "0")
                estable = oForm.Items.Item("Item_10").Specific
                ptoEmisor = oForm.Items.Item("Item_12").Specific
                estable.Value = "0".PadRight(3, "0")
                ptoEmisor.Value = "0".PadRight(3, "0")
                ofolder = oForm.Items.Item("Item_21").Specific
                ofolder.Select()
            Else
                oForm = Me.SBO_Application.Forms.Item("frm_inf")
                ruc = oForm.Items.Item("txtruc").Specific
                ruc.Value = "0".PadRight(13, "0")
                estable = oForm.Items.Item("Item_10").Specific
                ptoEmisor = oForm.Items.Item("Item_12").Specific
                estable.Value = "0".PadRight(3, "0")
                ptoEmisor.Value = "0".PadRight(3, "0")
            End If
            cargar()
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
            If (pVal.FormTypeEx = "60006" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = True) Then
                If pVal.ItemUID = "btnGuardar" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                    Dim comboA As SAPbouiCOM.ComboBox
                    Dim comboE As SAPbouiCOM.ComboBox
                    Dim identi As SAPbouiCOM.ComboBox
                    Dim conta As SAPbouiCOM.ComboBox
                    Dim razon As SAPbouiCOM.EditText
                    Dim nombre As SAPbouiCOM.EditText
                    Dim estable As SAPbouiCOM.EditText
                    Dim ptoEmisor As SAPbouiCOM.EditText
                    Dim direccion As SAPbouiCOM.EditText
                    Dim ci As SAPbouiCOM.EditText
                    Dim ruc As SAPbouiCOM.EditText
                    Dim dina As SAPbouiCOM.EditText
                    Dim rucct As SAPbouiCOM.EditText
                    Dim contri As SAPbouiCOM.EditText
                    Dim especial As SAPbouiCOM.EditText
                    
                    comboA = oForm.Items.Item("cboAmb").Specific
                    comboE = oForm.Items.Item("cboEmi").Specific
                    identi = oForm.Items.Item("Item_0").Specific
                    conta = oForm.Items.Item("Item_1").Specific
                    razon = oForm.Items.Item("Item_5").Specific
                    nombre = oForm.Items.Item("Item_7").Specific
                    estable = oForm.Items.Item("Item_10").Specific
                    ptoEmisor = oForm.Items.Item("Item_12").Specific
                    direccion = oForm.Items.Item("Item_14").Specific
                    ruc = oForm.Items.Item("txtruc").Specific
                    ci = oForm.Items.Item("txtCI").Specific
                    dina = oForm.Items.Item("Item_43").Specific
                    rucct = oForm.Items.Item("rucct").Specific
                    contri = oForm.Items.Item("contri").Specific
                    especial = oForm.Items.Item("numcte").Specific
                    If comboA.Value.Equals("") Then
                        Me.SBO_Application.SetStatusBarMessage("Debe de seleccionar un ambiente", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    End If
                    If comboE.Value.Equals("") Then
                        Me.SBO_Application.SetStatusBarMessage("Debe de seleccionar un Emisor", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    End If
                    If razon.Value.Equals("") Then
                        Me.SBO_Application.SetStatusBarMessage("Debe de seleccionar una Razon", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    End If
                    If nombre.Value.Equals("") Then
                        Me.SBO_Application.SetStatusBarMessage("Debe de seleccionar un Nombre Comercial", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    End If
                    If estable.Value.Equals("") Then
                        Me.SBO_Application.SetStatusBarMessage("Debe de seleccionar un establecimiento", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    Else
                        If estable.Value.ToString.Count <> 3 Then
                            Me.SBO_Application.SetStatusBarMessage("Establecimiento no valido, 3 digítos permitidos ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                            Return
                        End If
                       
                    End If
                    If ptoEmisor.Value.Equals("") Then
                        Me.SBO_Application.SetStatusBarMessage("Debe de seleccionar un Emisor", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    Else
                        If ptoEmisor.Value.ToString.Count <> 3 Then
                            Me.SBO_Application.SetStatusBarMessage("PtoEmisor no valido, 3 digítos permitidos ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                            Return
                        End If
                    End If
                    If direccion.Value.Equals("") Then
                        Me.SBO_Application.SetStatusBarMessage("Debe de seleccionar un direccion", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    End If
                    If ruc.Value.Equals("") Then
                        Me.SBO_Application.SetStatusBarMessage("Debe de escribir un RUC", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    Else
                        If ruc.Value.ToString.Count <> 13 Then
                            Me.SBO_Application.SetStatusBarMessage("RUC no válido, 13 digitos permitidos", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                            Return
                        Else
                            Try
                                Long.Parse(ruc.Value)
                            Catch ex As Exception
                                Me.SBO_Application.SetStatusBarMessage("RUC no válido", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                Return
                            End Try
                        End If
                    End If
                    If ci.Value = "" Then
                        Me.SBO_Application.SetStatusBarMessage("C.I. no válido", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    Else
                        If ci.Value.Count <> 10 Then
                            Me.SBO_Application.SetStatusBarMessage("C.I. no válido solo 10 digitos", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                            Return
                        Else
                            Try
                                Long.Parse(ci.Value)
                            Catch ex As Exception
                                Me.SBO_Application.SetStatusBarMessage("C.I. no válido solo dígitos", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                Return
                            End Try
                        End If
                    End If
                    If dina.Value = "" Then
                        Me.SBO_Application.SetStatusBarMessage("Debe de ingresar un Código DINARDAP", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    End If
                    If identi.Value.Trim = "" Then
                        Me.SBO_Application.SetStatusBarMessage("Debe de ingresar un tipo de identificación", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    End If

                    If rucct.Value.Equals("") Then
                        Me.SBO_Application.SetStatusBarMessage("Debe de escribir un RUC cliente", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    Else
                        If rucct.Value.ToString.Count <> 13 Then
                            Me.SBO_Application.SetStatusBarMessage("RUC cliente no válido, 13 digitos permitidos", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                            Return
                        Else
                            Try
                                Long.Parse(rucct.Value)
                            Catch ex As Exception
                                Me.SBO_Application.SetStatusBarMessage("RUC cliente no válido", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                Return
                            End Try
                        End If
                    End If

                    If conta.Value.Trim = "" Then
                        Me.SBO_Application.SetStatusBarMessage("Obligado a llevar contabilidad no seleccionado", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    End If

                    Dim orecord As SAPbobsCOM.Recordset
                    orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim sql As String = "Exec INSERTAR_INFOR_TRIBUTARIA " & comboA.Value & "," & comboE.Value & ",'" & razon.Value & "','" & nombre.Value & "','" & estable.Value & "','" & ptoEmisor.Value & "','" & direccion.Value & "','" & ruc.Value & "','" & ci.Value & "','" & dina.Value & "','" & identi.Value & "','" & rucct.Value & "','" & contri.Value & "','" & especial.Value & "','" & conta.Value & "','" & oCompany.CompanyName & "'"
                    orecord.DoQuery(sql)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                    orecord = Nothing
                    GC.Collect()
                    SBO_Application.SetStatusBarMessage("Informacion Guardada", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    BubbleEvent = False
                End If
            End If
        Catch ex As Exception
            Me.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, True)
        End Try
    End Sub

    Private Sub cargar()
        Dim orecord As SAPbobsCOM.Recordset
        orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            Dim comboA As SAPbouiCOM.ComboBox
            Dim comboE As SAPbouiCOM.ComboBox
            Dim razon As SAPbouiCOM.EditText
            Dim nombre As SAPbouiCOM.EditText
            Dim estable As SAPbouiCOM.EditText
            Dim ptoEmisor As SAPbouiCOM.EditText
            Dim direccion As SAPbouiCOM.EditText
            Dim ruc As SAPbouiCOM.EditText
            Dim identi As SAPbouiCOM.ComboBox
            Dim conta As SAPbouiCOM.ComboBox
            Dim ci As SAPbouiCOM.EditText
            Dim dina As SAPbouiCOM.EditText
            Dim rucct As SAPbouiCOM.EditText
            Dim contri As SAPbouiCOM.EditText
            Dim especial As SAPbouiCOM.EditText
            identi = oForm.Items.Item("Item_0").Specific
            conta = oForm.Items.Item("Item_1").Specific
            comboA = oForm.Items.Item("cboAmb").Specific
            comboE = oForm.Items.Item("cboEmi").Specific
            razon = oForm.Items.Item("Item_5").Specific
            nombre = oForm.Items.Item("Item_7").Specific
            estable = oForm.Items.Item("Item_10").Specific
            ptoEmisor = oForm.Items.Item("Item_12").Specific
            direccion = oForm.Items.Item("Item_14").Specific
            ruc = oForm.Items.Item("txtruc").Specific
            ci = oForm.Items.Item("txtCI").Specific
            dina = oForm.Items.Item("Item_43").Specific
            rucct = oForm.Items.Item("rucct").Specific
            contri = oForm.Items.Item("contri").Specific
            especial = oForm.Items.Item("numcte").Specific
            orecord.DoQuery("select * from [@INF_TRIBUTARIA]")
            If orecord.RecordCount > 0 Then
                While orecord.EoF = False
                    Dim valor = orecord.Fields.Item("U_AMBIENTE").Value
                    Dim valor2 = orecord.Fields.Item("U_EMISION").Value
                    comboA.Select(valor.ToString.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    comboE.Select(valor2.ToString.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    razon.Value = orecord.Fields.Item("U_RAZON_SOCIAL").Value
                    nombre.Value = orecord.Fields.Item("U_NOMBRE_COMERCIAL").Value
                    estable.Value = orecord.Fields.Item("U_ESTABLECIMIENTO").Value
                    ptoEmisor.Value = orecord.Fields.Item("U_PTO_EMISOR").Value
                    direccion.Value = orecord.Fields.Item("U_DIRECCION").Value
                    ruc.Value = orecord.Fields.Item("U_RUC").Value
                    ci.Value = orecord.Fields.Item("U_CI").Value
                    dina.Value = orecord.Fields.Item("U_COD_DINARDAP").Value
                    identi.Select(orecord.Fields.Item("U_TIP_IDENT").Value.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    rucct.Value = orecord.Fields.Item("U_RUC_CLIENTE").Value
                    contri.Value = orecord.Fields.Item("U_CLS_CONTRIBU").Value
                    especial.Value = orecord.Fields.Item("U_CLS_CONTRIBU_NUM").Value
                    conta.Select(orecord.Fields.Item("U_CONTA").Value.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    orecord.MoveNext()
                End While
            End If

        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
        System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
        orecord = Nothing
        GC.Collect()
    End Sub

End Class
