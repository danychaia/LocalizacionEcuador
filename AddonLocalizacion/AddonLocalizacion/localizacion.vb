'DANIEL MORENO
'ADDON LOCALIZACION ECUADOR
'ONESOLUTIONS
'MODULO DE ARRANQUE Y DEFINICION DE CAMPOS
'15/11/2016
Public Class localizacion
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCompany As SAPbobsCOM.Company
    Private oBusinessForm As SAPbouiCOM.Form
    Public oFilters As SAPbouiCOM.EventFilters
    Public oFilter As SAPbouiCOM.EventFilter
    Private oMatrix As SAPbouiCOM.Matrix        ' Global variable to handle matrixes

    ' Variables for Blanket Agreement UI form


    Private AddStarted As Boolean                ' Flag that indicates "Add" process started

    Private RedFlag As Boolean                   ' RedFlag when true indicates an error during "Add" process


#Region "Single Sign On"

    Private Sub SetApplication()

        AddStarted = False

        RedFlag = False

        '*******************************************************************

        '// Use an SboGuiApi object to establish connection

        '// with the SAP Business One application and return an

        '// initialized application object

        '*******************************************************************
        Try
            Dim SboGuiApi As SAPbouiCOM.SboGuiApi

            Dim sConnectionString As String

            SboGuiApi = New SAPbouiCOM.SboGuiApi

            '// by following the steps specified above, the following

            '// statement should be sufficient for either development or run mode
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            Else
                sConnectionString = Environment.GetCommandLineArgs.GetValue(0)
            End If

            'sConnectionString = Environment.GetCommandLineArgs.GetValue(1) '"0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

            '// connect to a running SBO Application

            SboGuiApi.Connect(sConnectionString)

            '// get an initialized application object

            SBO_Application = SboGuiApi.GetApplication()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try


    End Sub

    Private Function SetConnectionContext() As Integer

        Dim sCookie As String

        Dim sConnectionContext As String

        Dim lRetCode As Integer

        Try

            '// First initialize the Company object

            oCompany = New SAPbobsCOM.Company

            '// Acquire the connection context cookie from the DI API.

            sCookie = oCompany.GetContextCookie

            '// Retrieve the connection context string from the UI API using the

            '// acquired cookie.

            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)

            '// before setting the SBO Login Context make sure the company is not

            '// connected

            If oCompany.Connected = True Then

                oCompany.Disconnect()

            End If

            '// Set the connection context information to the DI API.

            SetConnectionContext = oCompany.SetSboLoginContext(sConnectionContext)

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Function

    Private Function ConnectToCompany() As Integer

        '// Establish the connection to the company database.

        ConnectToCompany = oCompany.Connect

    End Function

    Private Sub Class_Init()
        Try
            '//*************************************************************

            '// set SBO_Application with an initialized application object

            '//*************************************************************

            SetApplication()

            '//*************************************************************

            '// Set The Connection Context

            '//*************************************************************

            If Not SetConnectionContext() = 0 Then

                SBO_Application.MessageBox("Failed setting a connection to DI API")

                End ' Terminating the Add-On Application

            End If

            '//*************************************************************

            '// Connect To The Company Data Base

            '//*************************************************************

            If Not ConnectToCompany() = 0 Then

                SBO_Application.MessageBox("Failed connecting to the company's Data Base")

                End ' Terminating the Add-On Application

            End If

            '//*************************************************************

            '// send an "hello world" message

            '//*************************************************************

            SBO_Application.SetStatusBarMessage("DI Connected To: " & oCompany.CompanyName & vbNewLine & "Add-on is loaded", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            SetNewItems()
            'SetNewTax("01", "512 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "512", "_SYS00000000128")
            'SetNewTax("02", "513 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "513", "_SYS00000000128")
            'SetNewTax("03", "513A 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "513A", "_SYS00000000128")
            'SetNewTax("04", "514 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "514", "_SYS00000000128")

            UDT_UF.SBOApplication = Me.SBO_Application
            UDT_UF.Company = Me.oCompany
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

#End Region



    Public Sub New()

        MyBase.New()

        Class_Init()

        AddMenuItems()

        SetFilters()


    End Sub
    ''Function for add menus for SAP
    Private Sub AddMenuItems()

        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        oMenus = SBO_Application.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = (SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams))
        oMenuItem = SBO_Application.Menus.Item("43520") 'Modules
        If SBO_Application.Menus.Exists("localización") Then
            SBO_Application.Menus.RemoveEx("localización")
        End If
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
        oCreationPackage.UniqueID = "localización"
        oCreationPackage.String = "Localización"
        oCreationPackage.Enabled = True
        oCreationPackage.Position = 1
        oCreationPackage.Image = Application.StartupPath & "\locali.png"

        oMenus = oMenuItem.SubMenus

        Try
            'If the manu already exists this code will fail
            oMenus.AddEx(oCreationPackage)

            oMenuItem = SBO_Application.Menus.Item("localización")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "infTri"
            oCreationPackage.String = "Información Tributaria"
            oMenus.AddEx(oCreationPackage)

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "fact"
            oCreationPackage.String = "Comprobante Factura"
            oMenus.AddEx(oCreationPackage)

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "rete"
            oCreationPackage.String = "Comprobante para Retensión"
            oMenus.AddEx(oCreationPackage)

        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Long, True)
        End Try

    End Sub

    Private Sub SetFilters()

        '// Create a new EventFilters object

        oFilters = New SAPbouiCOM.EventFilters



        '// add an event type to the container

        '// this method returns an EventFilter object

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)



        '// assign the form type on which the event would be processed

        oFilter.AddEx("134") 'Quotation Form

        'oFilter.AddEx("139") 'Orders Form

        'oFilter.AddEx("133") 'Invoice Form

        'oFilter.AddEx("169") 'Main Menu



        SBO_Application.SetFilter(oFilters)

    End Sub


    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent

        Try
            If (pVal.ItemUID = "134" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = True) Then

            End If

            If pVal.FormTypeEx = "134" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.BeforeAction = True Then

                'oBusinessForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                ' Dim lab As SAPbouiCOM.StaticText
                'im label = oBusinessForm.Items.Add("lblRuc", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                'Dim panel = oBusinessForm.Items.Add("nuevo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                'label.Left = oBusinessForm.Left

                'label.Description = "RUC"
                'label.Left = 129
                'label.Top = 87
                'label.ToPane = 0
                'lab = label.Specific
                'lab.Caption = "nuevo"
                'panel.Top = 87
                'panel.Description = "nuevo "
                'panel.ToPane = 0
                'panel.Left = 139


            End If




        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        ' Events of the Blanket Agreement form
    End Sub


    Private Sub SBO_Application_DATAEVENT(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And pVal.FormTypeEx = "134" And pVal.BeforeAction = True Then
            Try
                Dim oform = SBO_Application.Forms.GetForm("134", 0)
                Dim oBPcode As SAPbouiCOM.EditText
                Dim oTipoIden As SAPbouiCOM.ComboBox
                Dim oUform = SBOApplication.Forms.GetForm("-134", 0)
                oTipoIden = oUform.Items.Item("U_IDENTIFICACION").Specific
                oBPcode = oform.Items.Item("5").Specific
                If oTipoIden.Value = "" Then
                    SBOApplication.SetStatusBarMessage("Debe de seleccionar un tipo de Identificación", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    BubbleEvent = False
                End If


                ''Cuando se selecciona un RUC
                If oTipoIden.Selected.Description = "RUC" And oBPcode.Value <> "" Then
                    Dim oDocumento As SAPbouiCOM.EditText
                    oDocumento = oUform.Items.Item("U_DOCUMENTO").Specific
                    oDocumento.Value = Trim(Right(oBPcode.Value, Len(oBPcode.Value) - 2)).ToString
                    If oDocumento.Value.ToString.Count = 13 Then
                        Try
                            Long.Parse(oDocumento.Value)
                        Catch ex As Exception
                            SBOApplication.SetStatusBarMessage("Para RUC solo se permiten Digitos", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                            BubbleEvent = False
                        End Try
                        Dim oTipoRuc As SAPbouiCOM.ComboBox
                        oTipoRuc = oUform.Items.Item("U_TIPO_RUC").Specific
                        If oTipoRuc.Value <> "" Then
                            If oTipoRuc.Selected.Description = "PUBLICO" Then
                                BubbleEvent = digitoVerificadorPublico(oDocumento.Value, SBOApplication)
                            Else
                                If oTipoRuc.Selected.Description = "NATURAL" Then
                                    BubbleEvent = digitoVerificadorIndividual(oDocumento.Value, SBOApplication)
                                    If oTipoRuc.Selected.Description = "PASAPORTE" Then

                                    End If
                                End If
                            End If
                        Else
                            SBOApplication.SetStatusBarMessage("Debe de seleccionar un tipo de RUC", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                            BubbleEvent = False
                        End If
                    Else
                        SBOApplication.SetStatusBarMessage("RUC debe contener 13 dígitos", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        BubbleEvent = False
                    End If
                Else
                    If oTipoIden.Selected.Description = "CEDULA" Then
                        Dim oDocumento As SAPbouiCOM.EditText
                        oDocumento = oUform.Items.Item("U_DOCUMENTO").Specific
                        oDocumento.Value = Trim(Right(oBPcode.Value, Len(oBPcode.Value) - 2)).ToString
                    Else
                        If oTipoIden.Selected.Description = "PASAPORTE" Then
                            Dim oDocumento As SAPbouiCOM.EditText
                            oDocumento = oUform.Items.Item("U_DOCUMENTO").Specific
                            If oDocumento.Value = "" Then
                                SBOApplication.SetStatusBarMessage("Debe de Ingresar un Pasaporte", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                BubbleEvent = False
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                SBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                BubbleEvent = False
            End Try


        End If

        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And pVal.FormTypeEx = "141" And pVal.BeforeAction = True Then
            Try
                Dim oUForm = SBOApplication.Forms.GetForm("-141", 0)
                Dim oSEstable As SAPbouiCOM.EditText
                Dim optoEmision As SAPbouiCOM.EditText
                Dim oEstableReten As SAPbouiCOM.EditText
                Dim optoRetencion As SAPbouiCOM.EditText
                Dim oSusTribu As SAPbouiCOM.ComboBox
                Dim oTipoComro As SAPbouiCOM.ComboBox                
                oSEstable = oUForm.Items.Item("U_SERIE_ESTABLE").Specific
                optoEmision = oUForm.Items.Item("U_PTO_EMISION").Specific
                oEstableReten = oUForm.Items.Item("U_STBLE_RETENCION").Specific
                optoRetencion = oUForm.Items.Item("U_PTO_RETENCION").Specific
                oSusTribu = oUForm.Items.Item("U_SUS_TRIBU").Specific
                oTipoComro = oUForm.Items.Item("U_TI_COMPRO").Specific
                If oSusTribu.Value = "" Then
                    SBOApplication.SetStatusBarMessage("Debe de seleccionar un sustento Tributario", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    BubbleEvent = False
                End If

                If oSEstable.Value.ToString.Count = 3 Then
                    Try
                        Integer.Parse(oSEstable.Value.ToString)
                    Catch ex As Exception
                        SBOApplication.SetStatusBarMessage("Serie de establecimiento permite dígitos", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        BubbleEvent = False
                    End Try

                Else
                    SBOApplication.SetStatusBarMessage("Serie de establecimiento debe de tener 3 digitos. ejemp 001", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    BubbleEvent = False
                End If


                If oSEstable.Value.ToString.Count = 3 Then
                    Try
                        Integer.Parse(oSEstable.Value.ToString)
                    Catch ex As Exception
                        SBOApplication.SetStatusBarMessage("Serie de establecimiento permite dígitos", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        BubbleEvent = False
                    End Try
                Else
                    SBOApplication.SetStatusBarMessage("Serie de establecimiento debe de tener 3 digitos. ejemp 001", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    BubbleEvent = False
                End If
            Catch ex As Exception
                SBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                BubbleEvent = False
            End Try


        End If

    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        If (pVal.MenuUID = "fact") And (pVal.BeforeAction = False) Then
            Dim fact As New fact_compro
            BubbleEvent = False
        End If
        If (pVal.MenuUID = "infTri") And (pVal.BeforeAction = False) Then
            Dim inf As New inf_tributaria
            BubbleEvent = False
        End If
        If (pVal.MenuUID = "rete") And (pVal.BeforeAction = False) Then
            Dim rete As New retencion
            BubbleEvent = False
        End If
    End Sub

    Private Sub SetNewItems()
        Try
            UDT_UF.userField(oCompany, "OCRD", "TIPO IDENTIFICACION", 45, "IDENTIFICACION", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBO_Application)
            UDT_UF.userField(oCompany, "OCRD", "TIPO RUC", 45, "TIPO_RUC", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "LOCAL O EXTERIOR", 45, "TIPO_CONTRI", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "NO. DOCUMENTO", 45, "DOCUMENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "OCRD", "RISE", 45, "RISE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "TIPO_SUJETO", 45, "TIPO_SUJETO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            'UDT_UF.userField(oCompany, "OCRD", "TIPO SUJETO", "TIPO_SUJETO", )
            UDT_UF.userTable(oCompany, "INF_TRIBUTARIA", "INFORMACION TRIBUTARIA", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "INF_PARTNER", "ADICIONAL AL PARTNER", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_MasterData, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_PARTNER", "DOBLE TRIBU", 30, "DO_TRI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "AMBIENTE", 11, "AMBIENTE", SAPbobsCOM.BoFieldTypes.db_Numeric, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "EMISION", 11, "EMISION", SAPbobsCOM.BoFieldTypes.db_Numeric, False, SBO_Application)
            'UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "EMISION", 11, "EMISION", SAPbobsCOM.BoFieldTypes.db_Numeric, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "RAZON SOCIAL", 250, "RAZON_SOCIAL", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "NOMBRE COMERCIAL", 250, "NOMBRE_COMERCIAL", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "ESTABLECIMIENTO", 45, "ESTABLECIMIENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "PTO EMISOR", 11, "PTO_EMISOR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "DIRECCION", 250, "DIRECCION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "RUC", 14, "RUC", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "CI", 45, "CI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "COD DINARDAP", 45, "COD_DINARDAP", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "TIPO IDENTI", 5, "TIP_IDENT", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "RUC CLIENTE", 14, "RUC_CLIENTE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "CLASE CONTRIBUYENTE", 45, "CLS_CONTRIBU", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "NO. CONTRIBUYENTE ESPECIAL", 45, "CLS_CONTRIBU_NUM", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "CONTA", 5, "CONTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "COMPANY", 55, "COMPANY", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)

            UDT_UF.userField(oCompany, "OINV", "ESTADO", 3, "ESTADO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OINV", "FIRMA", 60, "FIRMA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OWHT", "CODIGO ATS", 45, "COD_ATS", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "SUSTENTO TRIBUTARIO", 60, "SUS_TRIBU", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "TIPO COMPROBANTE", 45, "TI_COMPRO", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "SERIE DE ESTABLECIMIENTO", 3, "SERIE_ESTABLE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "PUNTO DE EMISION", 3, "PTO_EMISION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "SERIE DE ESTABLECIMIENTO", 3, "SERIE_ESTABLE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "PUNTO DE EMISION", 3, "PTO_EMISION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "ESTABLE. RETENCION", 3, "STBLE_RETENCION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "PUNTO DE RETENCION", 3, "PTO_RETENCION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "COMPROBANTE RETENCION", 45, "COMPRO_RETENCION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "AUTORIZACION RETENCION", 45, "AUTORI_RETENCION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "CADUCIDAD RETENCION", 45, "AUTORI_RETENCION", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "SUJETO A RETENCION", 25, "SUJE_RETENCION", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "PARTE RELACIONADA", 25, "PT_RELACIO", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "PUNTO DE EMISION", 3, "PTO_EMISION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            'UDT_UF.userField(oCompany, "OPCH", "SECUENCIAL", 3, "PTO_EMISION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            ''  updateValidValues()
        Catch ex As Exception
            ex.Message.ToString()
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub SetNewTax(wtCode As String, wtName As String, category As SAPbobsCOM.WithholdingTaxCodeCategoryEnum, baseType As SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum, baseAmount As Double, oficialCode As String, taxAccount As String)
        Try
            Dim erroS As String = " "
            Dim erro2 As Integer = 0
            Dim oTax As SAPbobsCOM.WithholdingTaxCodes
            oTax = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes)
            oTax.WTCode = wtCode
            oTax.WTName = wtName
            oTax.Category = category
            oTax.BaseType = baseType
            oTax.BaseAmount = baseAmount
            oTax.Lines.Effectivefrom = Date.Now
            oTax.Lines.Add()
            oTax.OfficialCode = oficialCode
            oTax.Account = taxAccount  ' "_SYS00000000128"
            oTax.UserFields.Fields.Item("U_COD_ATS").Value = "512"
            Dim recibe = oTax.Add()
            If recibe <> 0 Then
                oCompany.GetLastError(erro2, erroS)
                MessageBox.Show(erro2 & erroS)
            End If

        Catch ex As Exception
            SBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
    End Sub

    Private Function digitoVerificador(rucnum As String, application As SAPbouiCOM.Application) As Boolean
        Dim bandera As Boolean = True
        Dim provincia = rucnum.Chars(0) & rucnum.Chars(1)
        If provincia <= 0 And provincia >= 23 Then
            SBO_Application.SetStatusBarMessage("Error provincia no válida", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return bandera = False
        End If
        If rucnum.Chars(2) <> "9" Then
            application.SetStatusBarMessage("Error en el 3er Digito debe ser 9", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return bandera = False
        End If
        Dim pivote As Integer = 2
        Dim cantidadTotal As Integer = 0
        For i As Integer = 8 To 0 Step -1
            If pivote = 8 Then
                pivote = 2
            End If
            Dim temporal = Integer.Parse(rucnum.Chars(i))
            temporal *= pivote
            pivote += 1
            cantidadTotal += temporal
        Next
        cantidadTotal = 11 - (cantidadTotal Mod 11)
        If cantidadTotal.ToString = rucnum.Chars(9) Then
            Dim ultimos = rucnum.Chars(10) & rucnum.Chars(11) & rucnum.Chars(12)
            If ultimos = "000" Then
                application.SetStatusBarMessage("El numero de RUC no es válido en Principal o Sucursal ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                Return bandera = False
            Else
                application.SetStatusBarMessage("RUC válido", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            End If
        End If
        Return bandera = True
    End Function

    Private Function digitoVerificadorPublico(rucnum As String, application As SAPbouiCOM.Application) As Boolean
        Dim bandera As Boolean = True
        Dim provincia = rucnum.Chars(0) & rucnum.Chars(1)
        If provincia <= 0 And provincia >= 23 Then
            SBO_Application.SetStatusBarMessage("Error provincia no válida", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return bandera = False
        End If
        If rucnum.Chars(2) <> "6" Then
            application.SetStatusBarMessage("Error en el 3er Digito debe ser 6", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return bandera = False
        End If
        Dim pivote As Integer = 2
        Dim cantidadTotal As Integer = 0
        For i As Integer = 7 To 0 Step -1
            If pivote = 8 Then
                pivote = 2
            End If
            Dim temporal = Integer.Parse(rucnum.Chars(i))
            temporal *= pivote
            pivote += 1
            cantidadTotal += temporal
        Next
        cantidadTotal = 11 - (cantidadTotal Mod 11)
        If cantidadTotal.ToString = rucnum.Chars(8) Then
            Dim ultimos = rucnum.Chars(9) & rucnum.Chars(10) & rucnum.Chars(11) & rucnum.Chars(12)
            If ultimos = "0000" Then
                application.SetStatusBarMessage("El numero de RUC no es válido en Principal o Sucursal ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                Return bandera = False
            Else
                application.SetStatusBarMessage("RUC válido", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            End If
        Else
            application.SetStatusBarMessage("RUC no válido digito verficador no es corrrecto", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return bandera = False
        End If
        Return bandera = True
    End Function

    Private Function digitoVerificadorIndividual(rucnum As String, application As SAPbouiCOM.Application) As Boolean
        Dim bandera As Boolean = True
        Dim provincia = rucnum.Chars(0) & rucnum.Chars(1)
        If provincia <= 0 And provincia >= 23 Then
            SBO_Application.SetStatusBarMessage("Error provincia no válida", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return bandera = False
        End If
        If Integer.Parse(rucnum.Chars(2)) >= 1 And Integer.Parse(rucnum.Chars(2)) <= 5 Then
        Else
            application.SetStatusBarMessage("Error en el 3er Digito debe de estar en el rango de 1 a 5", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return bandera = False
        End If
        Dim pivote As Integer = 2
        Dim cantidadTotal As Integer = 0
        For i As Integer = 8 To 0 Step -1
            If pivote = 0 Then
                pivote = 2
            End If
            Dim temporal = Integer.Parse(rucnum.Chars(i))
            temporal *= pivote
            If temporal >= 10 Then
                Dim suma As Integer = 0
                For b As Integer = 0 To temporal.ToString.Count - 1 Step +1
                    suma += Integer.Parse(temporal.ToString.Chars(b))
                Next
                pivote -= 1
                cantidadTotal += suma
            Else
                pivote -= 1
                cantidadTotal += temporal
            End If

        Next
        cantidadTotal = 10 - (cantidadTotal Mod 10)
        If cantidadTotal.ToString = rucnum.Chars(9) Then
            Dim ultimos = rucnum.Chars(9) & rucnum.Chars(10) & rucnum.Chars(11) & rucnum.Chars(12)
            If ultimos = "0000" Then
                application.SetStatusBarMessage("El numero de RUC no es válido en Principal o Sucursal ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                Return bandera = False
            Else
                application.SetStatusBarMessage("RUC válido", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            End If
        End If
        Return bandera = True
    End Function

    Private Sub updateValidValues()
        Try
            Dim validArray As New ArrayList()
            Dim oValid As New validValues
            oValid.value = "04"
            oValid.descrip = "RUC"
            validArray.Add(oValid)

            oValid = New validValues
            oValid.value = "05"
            oValid.descrip = "CEDULA"
            validArray.Add(oValid)

            oValid = New validValues

            oValid.value = "06"
            oValid.descrip = "PASAPORTE"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "IDENTIFICACION", validArray)

            validArray.Clear()
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "PUBLICO"
            validArray.Add(oValid)

            oValid = New validValues
            oValid.value = "02"
            oValid.descrip = "NATURAL"
            validArray.Add(oValid)

            oValid = New validValues
            oValid.value = "03"
            oValid.descrip = "PASAPORTES"
            validArray.Add(oValid)

            UDT_UF.updateUserField(oCompany, "OCRD", "TIPO_RUC", validArray)

            validArray.Clear()
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "LOCAL"
            validArray.Add(oValid)

            oValid = New validValues
            oValid.value = "02"
            oValid.descrip = "EXTERNO"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "TIPO_CONTRI", validArray)

            validArray.Clear()
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "NATURAL"
            validArray.Add(oValid)

            oValid = New validValues
            oValid.value = "02"
            oValid.descrip = "NATURAL(RISE)"
            validArray.Add(oValid)

            oValid = New validValues
            oValid.value = "03"
            oValid.descrip = "SOCIEDADES"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "TIPO_SUJETO", validArray)

            validArray.Clear()
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "SI"
            validArray.Add(oValid)

            oValid = New validValues
            oValid.value = "02"
            oValid.descrip = "NO"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OPCH", "PT_RELACIO", validArray)

        Catch ex As Exception

        End Try

        
    End Sub

End Class
