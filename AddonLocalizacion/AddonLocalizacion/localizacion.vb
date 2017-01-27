Imports System.Xml
Imports System.Data.OleDb

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

            'SetNewTax("01", "512 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "512", "1-1-010-10-000")
            'SetNewTax("02", "513 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "513", "1-1-010-10-000")
            'SetNewTax("03", "513A 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "513A", "_SYS00000000128")
            'SetNewTax("04", "514 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "514", "_SYS00000000128")

            UDT_UF.SBOApplication = Me.SBO_Application
            UDT_UF.Company = Me.oCompany
            cargarInicial(oCompany, SBO_Application)
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


            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "CRtn"
            oCreationPackage.Position = "2"
            oCreationPackage.String = "Generar Comprobantes"
            oMenus.AddEx(oCreationPackage)


            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "Mta"
            oCreationPackage.Position = "3"
            oCreationPackage.String = "Mantenimiento"
            oMenus.AddEx(oCreationPackage)

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "inf"
            oCreationPackage.Position = "4"
            oCreationPackage.String = "Comprobantes Generados"
            oMenus.AddEx(oCreationPackage)


            ' MenuItem = SBO_Application.Menus.Item("CRtn")
            'oMenus = oMenuItem.SubMenus
            'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            'oCreationPackage.UniqueID = "fact"
            'oCreationPackage.String = "Comprobante Factura"
            'oMenus.AddEx(oCreationPackage)
            oMenuItem = SBO_Application.Menus.Item("CRtn")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "rete"
            oCreationPackage.String = "Comprobante para Retensión"
            oMenus.AddEx(oCreationPackage)



            oMenuItem = SBO_Application.Menus.Item("inf")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "rinfo"
            oCreationPackage.String = "Información de Retenciones de compras"
            oMenus.AddEx(oCreationPackage)

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "rinv"
            oCreationPackage.String = "Información de Retenciones de ventas"
            oMenus.AddEx(oCreationPackage)



            oMenuItem = SBO_Application.Menus.Item("Mta")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "CA"
            oCreationPackage.String = "Tipo Comprobante-Tipo Sustento"
            oMenus.AddEx(oCreationPackage)

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "ss"
            oCreationPackage.String = "Series"
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
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
        'oFilter = oFilter.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

        ' oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
        ' oFilter.AddEx("60006") 'Quotation Form



        '// assign the form type on which the event would be processed

        oFilter.AddEx("134") 'Quotation Form
        oFilter.AddEx("141")
        oFilter.AddEx("-141")
        oFilter.AddEx("133")
        oFilter.AddEx("60004")
        oFilter.AddEx("-133")
        oFilter.AddEx("-181")
        oFilter.AddEx("181")
        oFilter.AddEx("-65303")
        oFilter.AddEx("65303")
        oFilter.AddEx("65306")
        oFilter.AddEx("-65306")
        oFilter.AddEx("179")
        oFilter.AddEx("-179")
        'oFilter.AddEx("139") 'Orders Form
        'oFilter.AddEx("133") 'Invoice Form
        'oFilter.AddEx("169") 'Main Menu
        SBO_Application.SetFilter(oFilters)

    End Sub


    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent

        Try
            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                If pVal.FormTypeEx = "-141" And pVal.Before_Action = True And pVal.ItemUID = "U_RETENCION_NO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    Dim numero As New retencion_numeros
                    BubbleEvent = False
                    Return
                End If

                If pVal.FormTypeEx = "-133" And pVal.Before_Action = True And pVal.ItemUID = "U_RETENCION_NO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    Dim numero As New retencion_numeros
                    BubbleEvent = False
                    Return
                End If
                If pVal.FormTypeEx = "-65303" And pVal.Before_Action = True And pVal.ItemUID = "U_RETENCION_NO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    Dim numero As New retencion_numeros
                    BubbleEvent = False
                    Return
                End If
                If pVal.FormTypeEx = "-181" And pVal.Before_Action = True And pVal.ItemUID = "U_RETENCION_NO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    Dim numero As New retencion_numeros
                    BubbleEvent = False
                    Return
                End If

                If pVal.FormTypeEx = "-179" And pVal.Before_Action = True And pVal.ItemUID = "U_RETENCION_NO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    Dim numero As New retencion_numeros
                    BubbleEvent = False
                    Return
                End If
                If pVal.FormTypeEx = "-65306" And pVal.Before_Action = True And pVal.ItemUID = "U_RETENCION_NO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    Dim numero As New retencion_numeros
                    BubbleEvent = False
                    Return
                End If
                If pVal.FormTypeEx = "134" And pVal.Before_Action = True And pVal.ItemUID = "1" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        Dim oform = SBO_Application.Forms.Item(pVal.FormUID)
                        Dim oBPcode As SAPbouiCOM.EditText

                        Dim oTipoIden As SAPbouiCOM.ComboBox
                        Dim oTipoCliente As SAPbouiCOM.ComboBox
                        Dim oUform = SBOApplication.Forms.GetForm("-134", pVal.FormTypeCount)

                        oTipoIden = oUform.Items.Item("U_IDENTIFICACION").Specific
                        oBPcode = oform.Items.Item("5").Specific
                        oTipoCliente = oform.Items.Item("40").Specific
                        If oTipoCliente.Value = "C" Then
                            If oBPcode.Value.StartsWith("CN") = False And oBPcode.Value.StartsWith("CE") = False Then
                                SBOApplication.SetStatusBarMessage("El cliente debe de comenzar con CN o CE", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                System.Media.SystemSounds.Asterisk.Play()
                                BubbleEvent = False
                                Return
                            End If
                        Else
                            If oTipoCliente.Value = "S" Then
                                If oBPcode.Value.StartsWith("PL") = False And oBPcode.Value.StartsWith("PE") = False Then
                                    SBOApplication.SetStatusBarMessage("El cliente debe de comenzar con PL o PE", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    System.Media.SystemSounds.Asterisk.Play()
                                    BubbleEvent = False
                                    Return
                                End If
                            End If
                        End If

                        If oTipoIden.Value = "" Then
                            SBOApplication.SetStatusBarMessage("Debe de seleccionar un tipo de Identificación", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            System.Media.SystemSounds.Asterisk.Play()
                            BubbleEvent = False
                            Return
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
                                    SBOApplication.SetStatusBarMessage("Para RUC solo se permiten Digitos", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    System.Media.SystemSounds.Asterisk.Play()
                                    BubbleEvent = False
                                    Return
                                End Try
                                Dim oTipoRuc As SAPbouiCOM.ComboBox
                                oTipoRuc = oUform.Items.Item("U_TIPO_RUC").Specific
                                If oTipoRuc.Value <> "" Then
                                    If oTipoRuc.Selected.Description = "PUBLICO" Then
                                        BubbleEvent = digitoVerificadorPublico(oDocumento.Value, SBOApplication, True)
                                    Else
                                        If oTipoRuc.Selected.Description = "NATURAL" Then
                                            BubbleEvent = digitoVerificadorIndividual(oDocumento.Value, SBOApplication, False)
                                        Else
                                            If oTipoRuc.Selected.Description = "EXTRANJERO" Then
                                                BubbleEvent = digitoVerificador(oDocumento.Value, Me.SBO_Application, True)
                                            End If
                                        End If
                                    End If
                                Else
                                    SBOApplication.SetStatusBarMessage("Debe de seleccionar un tipo de RUC", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    System.Media.SystemSounds.Asterisk.Play()
                                    BubbleEvent = False
                                End If
                            Else
                                SBOApplication.SetStatusBarMessage("RUC debe contener 13 dígitos", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                System.Media.SystemSounds.Asterisk.Play()
                                BubbleEvent = False
                            End If
                        Else
                            If oTipoIden.Selected.Description = "CEDULA" Then
                                Dim oDocumento As SAPbouiCOM.EditText
                                oDocumento = oUform.Items.Item("U_DOCUMENTO").Specific
                                oDocumento.Value = Trim(Right(oBPcode.Value, Len(oBPcode.Value) - 2)).ToString
                                If oDocumento.Value.Count <> 10 Then
                                    SBOApplication.SetStatusBarMessage("Para Cedula se permiten solamente 10 dígitos", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    System.Media.SystemSounds.Asterisk.Play()
                                    BubbleEvent = False
                                    Return
                                Else
                                    Try
                                        Long.Parse(oDocumento.Value)
                                    Catch ex As Exception
                                        SBOApplication.SetStatusBarMessage("Para cedula no se permiten caracteres.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        System.Media.SystemSounds.Asterisk.Play()
                                        BubbleEvent = False
                                        Return
                                    End Try
                                    Dim oTipoRuc As SAPbouiCOM.ComboBox
                                    oTipoRuc = oUform.Items.Item("U_TIPO_RUC").Specific

                                    If oTipoRuc.Value <> "" Then
                                        If oTipoRuc.Selected.Description = "PUBLICO" Then
                                            BubbleEvent = digitoVerificadorPublico(oDocumento.Value, SBOApplication, True)
                                        Else
                                            If oTipoRuc.Selected.Description = "NATURAL" Then
                                                BubbleEvent = digitoVerificadorIndividual(oDocumento.Value, SBOApplication, True)
                                            Else

                                                If oTipoRuc.Selected.Description = "EXTRANJERO" Then
                                                    SBO_Application.SetStatusBarMessage("Para Cedula no se puede seleccionar un tipo RUC EXTRANJERO", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                    System.Media.SystemSounds.Asterisk.Play()
                                                    BubbleEvent = False
                                                    Return
                                                End If
                                            End If
                                        End If
                                    Else
                                        SBOApplication.SetStatusBarMessage("Debe un Tipo de Cédula", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        System.Media.SystemSounds.Asterisk.Play()
                                        BubbleEvent = False
                                        Return
                                    End If

                                End If
                            Else
                                If oTipoIden.Selected.Description = "PASAPORTE" Then
                                    Dim oDocumento As SAPbouiCOM.EditText
                                    oDocumento = oUform.Items.Item("U_DOCUMENTO").Specific
                                    If oDocumento.Value = "" Then
                                        SBOApplication.SetStatusBarMessage("Debe de Ingresar un Pasaporte", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        System.Media.SystemSounds.Asterisk.Play()
                                        BubbleEvent = False
                                        Return
                                    Else
                                        If oDocumento.Value <> "" Then
                                            Dim resp = SBO_Application.MessageBox("Guardara el documento con NO." & oDocumento.Value.Trim, 1, "SI.", "NO.")
                                            If resp = 2 Then
                                                SBOApplication.SetStatusBarMessage("Debe de Ingresar un Pasaporte", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                oDocumento.Value = ""
                                                System.Media.SystemSounds.Asterisk.Play()
                                                BubbleEvent = False
                                                Return
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If


                If pVal.FormTypeEx = "141" And pVal.Before_Action = True And pVal.ItemUID = "1" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        Try
                            Dim oUForm = SBOApplication.Forms.GetForm("-141", pVal.FormTypeCount)
                            Dim oNumRetencion As SAPbouiCOM.EditText
                            ' Dim oSEstable As SAPbouiCOM.EditText
                            ' Dim optoEmision As SAPbouiCOM.EditText
                            'Dim oEstableReten As SAPbouiCOM.EditText
                            ' Dim optoRetencion As SAPbouiCOM.EditText
                            Dim oSusTribu As SAPbouiCOM.ComboBox
                            Dim oTipoComro As SAPbouiCOM.ComboBox
                            'Dim oAutoRetencion As SAPbouiCOM.EditText

                            ' oSEstable = oUForm.Items.Item("U_SERIE_ESTABLE").Specific
                            'optoEmision = oUForm.Items.Item("U_PTO_EMISION").Specific
                            'oEstableReten = oUForm.Items.Item("U_STBLE_RETENCION").Specific
                            'optoRetencion = oUForm.Items.Item("U_PTO_RETENCION").Specific
                            oSusTribu = oUForm.Items.Item("U_SUS_TRIBU").Specific
                            oTipoComro = oUForm.Items.Item("U_TI_COMPRO").Specific
                            oNumRetencion = oUForm.Items.Item("U_RETENCION_NO").Specific
                            If UDT_UF.code <> "" Then
                                oNumRetencion.Value = UDT_UF.code
                            End If

                            ' oAutoRetencion = oUForm.Items.Item("U_AUTORI_RETENCION").Specific



                            ' Else
                            '       SBOApplication.SetStatusBarMessage("Punto de emisión establecimiento debe de tener 3 digitos. ejemp 001", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            '    BubbleEvent = False
                            '    Return
                            ' End If

                            ' If oEstableReten.Value.ToString.Count = 3 Then
                            'Try
                            'Integer.Parse(oEstableReten.Value.ToString)
                            ' Catch ex As Exception
                            '  SBOApplication.SetStatusBarMessage("Establecimiento de retención 3 permite dígitos", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            ' BubbleEvent = False
                            ' Return
                            '   End Try
                            '  Else
                            '   SBOApplication.SetStatusBarMessage("Establecimiento de retención debe de tener 3 digitos. ejemp 001", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            '   BubbleEvent = False
                            ' Return
                            ' End If
                        Catch ex As Exception
                            SBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            BubbleEvent = False
                            Return
                        End Try
                    End If
                End If

                If pVal.FormTypeEx = "133" And pVal.Before_Action = True And pVal.ItemUID = "1" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                        Try
                            Dim oAutoRetencion As SAPbouiCOM.EditText
                            Dim oUForm = SBOApplication.Forms.GetForm("-133", pVal.FormTypeCount)
                            oAutoRetencion = oUForm.Items.Item("U_RETENCION_NO").Specific
                            oAutoRetencion.Value = UDT_UF.code
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    End If
                End If


                If pVal.FormTypeEx = "181" And pVal.Before_Action = True And pVal.ItemUID = "1" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                        Try
                            Dim oAutoRetencion As SAPbouiCOM.EditText
                            Dim oUForm = SBOApplication.Forms.GetForm("-181", pVal.FormTypeCount)
                            oAutoRetencion = oUForm.Items.Item("U_RETENCION_NO").Specific
                            oAutoRetencion.Value = UDT_UF.code
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    End If
                End If

                If pVal.FormTypeEx = "65303" And pVal.Before_Action = True And pVal.ItemUID = "1" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                        Try
                            Dim oAutoRetencion As SAPbouiCOM.EditText
                            Dim oUForm = SBOApplication.Forms.GetForm("-65303", pVal.FormTypeCount)
                            oAutoRetencion = oUForm.Items.Item("U_RETENCION_NO").Specific
                            oAutoRetencion.Value = UDT_UF.code
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    End If
                End If

                If pVal.FormTypeEx = "179" And pVal.Before_Action = True And pVal.ItemUID = "1" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                        Try
                            Dim oAutoRetencion As SAPbouiCOM.EditText
                            Dim oUForm = SBOApplication.Forms.GetForm("-179", pVal.FormTypeCount)
                            oAutoRetencion = oUForm.Items.Item("U_RETENCION_NO").Specific
                            oAutoRetencion.Value = UDT_UF.code
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    End If
                End If
                If pVal.FormTypeEx = "65306" And pVal.Before_Action = True And pVal.ItemUID = "1" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                        Try
                            Dim oAutoRetencion As SAPbouiCOM.EditText
                            Dim oUForm = SBOApplication.Forms.GetForm("-65306", pVal.FormTypeCount)
                            oAutoRetencion = oUForm.Items.Item("U_RETENCION_NO").Specific
                            oAutoRetencion.Value = UDT_UF.code
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    End If
                End If
            End If 
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        ' Events of the Blanket Agreement form
    End Sub


    Private Sub SBO_Application_DATAEVENT(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try

            If pVal.FormTypeEx = "141" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                UDT_UF.code = ""
                Dim doc As New XmlDocument
                doc.LoadXml(pVal.ObjectKey)
                Dim docEntrynode = doc.DocumentElement.SelectSingleNode("/DocumentParams/DocEntry")
                Dim xmlRetencion As New generarRetencionXML
                xmlRetencion.generaRetencionXML(docEntrynode.InnerText, "compra", SBOApplication, oCompany)
            End If
            If pVal.FormTypeEx = "133" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                'MessageBox.Show(pVal.ObjectKey)
                Dim doc As New XmlDocument
                doc.LoadXml(pVal.ObjectKey)
                Dim docEntrynode = doc.DocumentElement.SelectSingleNode("/DocumentParams/DocEntry")
                generarXML(docEntrynode.InnerText, "13")
            End If
        Catch ex As Exception
            SBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try


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
        If (pVal.MenuUID = "rinfo") And (pVal.BeforeAction = False) Then
            Dim info_retencion As New retencion_info
            BubbleEvent = False
        End If
        If (pVal.MenuUID = "CA") And (pVal.BeforeAction = False) Then
            Dim ca As New comprobantes_autorizados
            BubbleEvent = False
        End If
        If (pVal.MenuUID = "ss") And (pVal.BeforeAction = False) Then
            Dim series As New series
            BubbleEvent = False
        End If
        If (pVal.MenuUID = "rinv") And (pVal.BeforeAction = False) Then
            Dim ventas As New retencion_info_ventas
            BubbleEvent = False
        End If
    End Sub

    Private Sub SetNewItems()
        Try
            UDT_UF.userField(oCompany, "OCRD", "TIPO IDENTIFICACION", 45, "IDENTIFICACION", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBO_Application)
            UDT_UF.userField(oCompany, "OCRD", "TIPO RUC", 45, "TIPO_RUC", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "LOCAL O EXTERIOR", 45, "TIPO_CONTRI", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "NO. DOCUMENTO", 45, "DOCUMENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            ' UDT_UF.userField(oCompany, "OCRD", "RISE", 45, "RISE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "TIPO_SUJETO", 45, "TIPO_SUJETO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "PAIS PAGO", 45, "PAIS_PAGO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "FORMA DE PAGO", 45, "FORMA_PAGO", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            'UDT_UF.userField(oCompany, "OCRD", "TIPO SUJETO", "TIPO_SUJETO", )
            UDT_UF.userTable(oCompany, "INF_TRIBUTARIA", "INFORMACION TRIBUTARIA", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "INF_PARTNER", "ADICIONAL AL PARTNER", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_MasterData, False, SBOApplication)
            UDT_UF.userTable(oCompany, "PAIS", "REGISTRO DE PAIS", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "MUNI_CANTO", "CANTON O MUNICIPIO", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "PARROQUIAS", "PARROQUIAS", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userField(oCompany, "@PARROQUIAS", "CANTON", 30, "CANTON", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@PARROQUIAS", "PRIVINCIA", 30, "PROVINCIA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

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
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "NUMERO DE ESTABLECIMIENTO", 6, "NO_ESTABLE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "TIPO DE SISTEMA", 6, "T_SISTEMA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)

            UDT_UF.userTable(oCompany, "COMPRO_AUTO", "INFORMACION AUTORIZACIONES", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userField(oCompany, "@COMPRO_AUTO", "CODIGO DE AUTORIZACION", 8, "C_CODE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@COMPRO_AUTO", "TIPO COMPROBANTE", 45, "TIPO_COMPRO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@COMPRO_AUTO", "SUSTENTO TRIBUTARIO", 25, "CODE_SUSTENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userTable(oCompany, "SERIES", "INFORMACION AUTORIZACIONES", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "CODIGO ESTABLECIMIENTO", 3, "E_CODE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "CODIGO PUNTO EMISION", 3, "P_CODE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "FECHA", 25, "DATE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "CORRELATIVO", 25, "CORRELATIVO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "TIPO DOCUMENTO", 45, "T_DOCUMENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "DE", 45, "DE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "HASTA", 45, "HASTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "NO AUTORIZACION", 45, "NO_AUTORI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "FECHA CADUCIDAD", 45, "FECHA_CADU", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "DOCUMENTO", 70, "DOCUMENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userField(oCompany, "OINV", "ESTADO", 3, "ESTADO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OINV", "NO. AUTORIZACION", 60, "NO_AUTORI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userField(oCompany, "OWHT", "CODIGO ATS", 45, "COD_ATS", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userField(oCompany, "OPCH", "SUSTENTO TRIBUTARIO", 60, "SUS_TRIBU", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "TIPO COMPROBANTE", 45, "TI_COMPRO", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)

            UDT_UF.userField(oCompany, "OPCH", "SUJETO A RETENCION", 25, "SUJE_RETENCION", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "PARTE RELACIONADA", 25, "PT_RELACIO", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "No. RETENCION", 45, "RETENCION_NO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "MONTO ICE ", 11, "MONTO_ICE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "ID PROVEEDOR ", 13, "ID_PROVEEDOR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "ID PROVEEDOR ", 13, "ID_PROVEEDOR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "MONTO IVA ", 11, "MONTO_IVA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "AUTORIZACION REEMBOLSO ", 11, "AUTO_REEMBOLSO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "MONTO ICE ", 11, "MONTO_ICE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "ORPC", "No. RETENCION", 45, "RETENCION_NO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            'UDT_UF.userField(oCompany, "OPCH", "SECUENCIAL", 3, "PTO_EMISION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            ''  updateValidValues()
        Catch ex As Exception
            ex.Message.ToString()
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub SetNewTax(wtCode As String, wtName As String, category As SAPbobsCOM.WithholdingTaxCodeCategoryEnum, baseType As SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum, baseAmount As Double, oficialCode As String, taxAccount As String, ATSCode As String, rate As Double)
        Try

            Dim erroS As String = " "
            Dim erro2 As Integer = 0
            Dim oTax As SAPbobsCOM.WithholdingTaxCodes
            oTax = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes)
            If oTax.GetByKey(wtCode) = False Then
                oTax.WTCode = wtCode
                oTax.WTName = wtName
                oTax.WithholdingType = SAPbobsCOM.WithholdingTypeEnum.wt_IncomeTaxWithholding
                oTax.Category = category
                oTax.BaseType = baseType
                oTax.BaseAmount = baseAmount
                oTax.Lines.Effectivefrom = Date.Now
                oTax.Lines.Rate = rate
                oTax.Lines.Add()
                oTax.OfficialCode = oficialCode
                oTax.Account = taxAccount  ' "_SYS00000000128"
                oTax.UserFields.Fields.Item("U_COD_ATS").Value = ATSCode
                Dim recibe = oTax.Add()
                If recibe <> 0 Then
                    oCompany.GetLastError(erro2, erroS)
                    MessageBox.Show(erro2 & erroS)
                End If
            End If


        Catch ex As Exception
            SBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
    End Sub

    Private Function digitoVerificador(rucnum As String, application As SAPbouiCOM.Application, cedula As Boolean) As Boolean
        Dim bandera As Boolean = True
        Dim provincia = rucnum.Chars(0) & rucnum.Chars(1)
        If provincia >= 0 Then
            If provincia <= 22 Then
            Else
                SBO_Application.SetStatusBarMessage("Error provincia no válida", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                Return bandera = False
            End If
        Else
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
        If (cantidadTotal Mod 11) = 0 Then
            cantidadTotal = 0
        Else
            cantidadTotal = 11 - (cantidadTotal Mod 11)
        End If
        If cantidadTotal.ToString = rucnum.Chars(9) Then

            If rucnum.EndsWith("001") = False Then
                application.SetStatusBarMessage("El numero de RUC no es válido en Principal o Sucursal ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                Return bandera = False
            Else
                application.SetStatusBarMessage("RUC válido", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            End If
        Else
            application.SetStatusBarMessage("El numero de RUC no es válido para el Dígito Verificador ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return bandera = False
        End If
        Return bandera = True
    End Function

    Private Function digitoVerificadorPublico(rucnum As String, application As SAPbouiCOM.Application, cedula As Boolean) As Boolean
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
        If (cantidadTotal Mod 11) = 0 Then
            cantidadTotal = 0
        Else
            cantidadTotal = 11 - (cantidadTotal Mod 11)
        End If
        If cantidadTotal.ToString = rucnum.Chars(8) Then
            If cedula = False Then
                If rucnum.EndsWith("001") = False Then
                    application.SetStatusBarMessage("El numero de RUC no es válido en Principal o Sucursal ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    Return bandera = False
                End If

            Else
                application.SetStatusBarMessage("RUC válido", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            End If
        Else
            application.SetStatusBarMessage("RUC no válido digito verficador no es corrrecto", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return bandera = False
        End If
        Return bandera = True
    End Function

    Private Function digitoVerificadorIndividual(rucnum As String, application As SAPbouiCOM.Application, cedula As Boolean) As Boolean
        Dim bandera As Boolean = True
        Dim provincia = rucnum.Chars(0) & rucnum.Chars(1)
        If provincia >= 0 Then
            If provincia <= 22 Then
            Else
                SBO_Application.SetStatusBarMessage("Error provincia no válida", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                Return bandera = False
            End If
        Else
            SBO_Application.SetStatusBarMessage("Error provincia no válida", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return bandera = False
        End If
        If Integer.Parse(rucnum.Chars(2)) >= 0 And Integer.Parse(rucnum.Chars(2)) <= 5 Then
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
        If (cantidadTotal Mod 10) = 0 Then
            cantidadTotal = 0
        Else
            cantidadTotal = 10 - (cantidadTotal Mod 10)
        End If
        If cantidadTotal.ToString = rucnum.Chars(9) Then
            If cedula = False Then
                If rucnum.EndsWith("001") = False Then
                    application.SetStatusBarMessage("El numero de RUC no es válido en Principal o Sucursal ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    Return bandera = False
                End If

            Else
                'application.SetStatusBarMessage("RUC válido", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            End If
        Else
            application.SetStatusBarMessage("El dígito verificador es incorrecto ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return bandera = False
        End If
        Return bandera = True
    End Function

    Private Sub updateValidValues()
        Try
            Dim tabla As String
            Dim campo As String
            Dim validArrayList As New ArrayList()
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\Sustentos.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()
                    If currentLine(0).ToString = "SUSTENTO" And currentLine(1).ToString = "OPCH" Then
                        tabla = currentLine(1).ToString
                        campo = currentLine(2).ToString
                    Else
                        Dim oValidV As New validValues
                        oValidV.value = currentLine(0).ToString
                        oValidV.descrip = currentLine(1).ToString
                        validArrayList.Add(oValidV)
                        If currentLine(3).ToString = "fin" Then
                            UDT_UF.updateUserField(oCompany, tabla, campo, validArrayList)
                            validArrayList.Clear()
                        End If
                    End If
                End While
            End Using

            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\Comprobantes.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()
                    If currentLine(0).ToString = "Comprobantes" And currentLine(1).ToString = "OPCH" Then
                        tabla = currentLine(1).ToString
                        campo = currentLine(2).ToString
                    Else
                        Dim oValidV As New validValues
                        oValidV.value = currentLine(0).ToString
                        oValidV.descrip = currentLine(1).ToString
                        validArrayList.Add(oValidV)
                        If currentLine(3).ToString = "fin" Then
                            UDT_UF.updateUserField(oCompany, tabla, campo, validArrayList)
                            validArrayList.Clear()
                        End If
                    End If
                End While
            End Using


            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\Cantones.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()
                    If currentLine(0).ToString <> "[@MUNI_CANTO]" Then
                        Dim oRecord As SAPbobsCOM.Recordset
                        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim SQL = "INSERT INTO [@MUNI_CANTO] VALUES ('" & currentLine(0) & "','" & currentLine(1).ToString & "')"
                        oRecord.DoQuery(SQL)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                        oRecord = Nothing
                        GC.Collect()
                    End If
                End While
            End Using


            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\pais.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()
                    If currentLine(0).ToString <> "[@PAIS]" Then
                        Dim oRecord As SAPbobsCOM.Recordset
                        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim SQL = "INSERT INTO [@PAIS] VALUES ('" & currentLine(1) & "','" & currentLine(0).ToString & "')"
                        oRecord.DoQuery(SQL)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                        oRecord = Nothing
                        GC.Collect()
                    End If
                End While
            End Using


            Dim validArray As New ArrayList()
            Dim oValid As New validValues
            oValid.value = "04"
            oValid.descrip = "RUC"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "05"
            oValid.descrip = "CEDULA"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "06"
            oValid.descrip = "PASAPORTE"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "IDENTIFICACION", validArray)

            validArray.Clear()
            oValid = Nothing
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "PUBLICO"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "02"
            oValid.descrip = "NATURAL"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "03"
            oValid.descrip = "PASAPORTES"
            validArray.Add(oValid)

            UDT_UF.updateUserField(oCompany, "OCRD", "TIPO_RUC", validArray)

            validArray.Clear()
            oValid = Nothing
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "LOCAL"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "02"
            oValid.descrip = "EXTERNO"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "TIPO_CONTRI", validArray)

            validArray.Clear()
            oValid = Nothing
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "NATURAL"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "02"
            oValid.descrip = "NATURAL(RISE)"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues
            oValid.value = "03"
            oValid.descrip = "SOCIEDADES"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "TIPO_SUJETO", validArray)

            validArray.Clear()
            oValid = Nothing
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "SI"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "02"
            oValid.descrip = "NO"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OPCH", "PT_RELACIO", validArray)
            validArray.Clear()

            oValid = Nothing
            oValid = New validValues

            oValid.value = "01"
            oValid.descrip = "SIN UTILIZACION DEL SISTEMA FINANCIERO"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "02"
            oValid.descrip = "CHEQUE PROPIO"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "03"
            oValid.descrip = "CHEQUE CERTIFICADO"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "04"
            oValid.descrip = "CHEQUE DE GERENCIA"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "05"
            oValid.descrip = "CHEQUE DEL EXTERIOR"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues
            oValid.value = "06"
            oValid.descrip = "DÉBITO DE CUENTA"
            validArray.Add(oValid)

            UDT_UF.updateUserField(oCompany, "OCRD", "FORMA_PAGO", validArray)
            validArray.Clear()


            oValid = Nothing
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "SI"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "02"
            oValid.descrip = "NO"
            validArray.Add(oValid)

            UDT_UF.updateUserField(oCompany, "OPCH", "SUJE_RETENCION", validArray)
            validArray.Clear()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub


    Private Function fieldExist(oCompany As SAPbobsCOM.Company, tableName As String, namefield As String) As Boolean

        Dim existe As Boolean = False
        Dim record As SAPbobsCOM.Recordset

        record = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        record.DoQuery("SELECT a.AliasID   FROM CUFD a WHERE TableID = '" & tableName & "' AND AliasID = '" & namefield & "'")
        If record.RecordCount > 0 Then
            existe = True
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(record)
        record = Nothing
        GC.Collect()
        Return existe
    End Function
    Private Sub generarXML(DocEntry As String, objectType As String)

        Dim doc As New XmlDocument
        Dim oRecord As SAPbobsCOM.Recordset
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oRecord.DoQuery("exec ENCABEZADO_FACTURA '" & DocEntry & "'")
        Dim writer As New XmlTextWriter("Comprobante (F) No." & DocEntry.ToString & ".xml", System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("factura")
        writer.WriteAttributeString("id", "comprobante")
        writer.WriteAttributeString("version", "2.0.0")
        writer.WriteStartElement("infoTributaria")
        createNode("ambiente", oRecord.Fields.Item(0).Value.ToString, writer)
        createNode("tipoEmision", oRecord.Fields.Item(1).Value.ToString, writer)
        createNode("razonSocial", oRecord.Fields.Item(2).Value.ToString, writer)
        createNode("ruc", oRecord.Fields.Item(3).Value.ToString.PadLeft(13, "0"), writer)
        'createNode("claveAcesso", claveAcceso(oRecord).PadLeft(49, "0"), writer)
        createNode("claveAcesso", "", writer)
        createNode("codDoc", oRecord.Fields.Item("codDoc").Value.ToString.PadLeft(2, "0"), writer)
        createNode("estab", oRecord.Fields.Item("estable").Value.ToString.PadLeft(3, "0"), writer)
        createNode("ptoEmi", oRecord.Fields.Item("ptoemi").Value.ToString.PadLeft(3, "0"), writer)
        createNode("secuencial", oRecord.Fields.Item("secuencial").Value.ToString.PadLeft(9, "0"), writer)
        createNode("dirMatriz", oRecord.Fields.Item("dirMatriz").Value.ToString, writer)
        ''Cierre info Tributaria
        writer.WriteEndElement()

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
        oRecord = Nothing
        GC.Collect()

        writer.WriteStartElement("infoFactura")
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecord.DoQuery("exec SP_INFO_FACTURA '" & DocEntry & "'")
        createNode("fechaEmision", Date.Parse(oRecord.Fields.Item("DATE").Value.ToString).ToString("dd/MM/yyyy"), writer)
        createNode("tipoIdentificacionComprador", oRecord.Fields.Item("U_IDENTIFICACION").Value.ToString, writer)
        createNode("razonSocialComprador", oRecord.Fields.Item("CardName").Value.ToString, writer)
        createNode("identificacionComprador", oRecord.Fields.Item("U_DOCUMENTO").Value.ToString, writer)
        createNode("totalSinImpuestos", oRecord.Fields.Item("sin_impuesto").Value.ToString, writer)
        createNode("totalDescuento", oRecord.Fields.Item("totDescuento").Value.ToString, writer)

        writer.WriteStartElement("totalConImpuestos")
        Dim importeTotal = oRecord.Fields.Item("DocTotal").Value.ToString

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
        oRecord = Nothing
        GC.Collect()

        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecord.DoQuery("exec SP_Total_Con_Impuesto '" & DocEntry & "'")
        If oRecord.RecordCount > 0 Then
            While oRecord.EoF = False
                writer.WriteStartElement("totalImpuesto")
                createNode("codigo", oRecord.Fields.Item(0).Value.ToString, writer)
                createNode("codigoPorcentaje", oRecord.Fields.Item(1).Value.ToString, writer)
                createNode("baseImponible", oRecord.Fields.Item(2).Value.ToString, writer)
                createNode("valor", oRecord.Fields.Item(2).Value.ToString, writer)
                createNode("baseNoGraIva", Double.Parse("0.00").ToString, writer)
                createNode("baseImponible", "", writer)
                writer.WriteEndElement()
                oRecord.MoveNext()
            End While
        End If


        ''Cierre TotalConImpuestos
        writer.WriteEndElement()
        createNode("importeTotal", importeTotal, writer)
        ''Cierre INFO FACTURA
        writer.WriteEndElement()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
        oRecord = Nothing
        GC.Collect()

        writer.WriteStartElement("detalles")
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecord.DoQuery("exec sp_DetalleFac '" & DocEntry & "'")


        If oRecord.RecordCount > 0 Then

            While oRecord.EoF = False
                Dim oRecord2 As SAPbobsCOM.Recordset
                oRecord2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                writer.WriteStartElement("detalle")
                createNode("descripcion", oRecord.Fields.Item(0).Value.ToString, writer)
                createNode("cantidad", oRecord.Fields.Item(1).Value.ToString, writer)
                createNode("precioUnitario", oRecord.Fields.Item(2).Value.ToString, writer)
                createNode("descuento", oRecord.Fields.Item(3).Value.ToString, writer)
                createNode("precioTotalSinImpuesto", oRecord.Fields.Item(4).Value.ToString, writer)
                writer.WriteStartElement("impuestos")
                oRecord2.DoQuery("exec SP_Impuesto_Detalle '" & DocEntry & "','" & oRecord.Fields.Item(5).Value.ToString & "'")
                If oRecord2.RecordCount > 0 Then
                    While oRecord2.EoF = False
                        writer.WriteStartElement("impuesto")
                        createNode("codigo", oRecord2.Fields.Item(0).Value.ToString, writer)
                        createNode("codigoPorcentaje", oRecord2.Fields.Item(1).Value.ToString, writer)
                        createNode("tarifa", oRecord2.Fields.Item(3).Value.ToString, writer)
                        createNode("baseImponible", oRecord2.Fields.Item(2).Value.ToString, writer)
                        createNode("valor", oRecord2.Fields.Item(2).Value.ToString, writer)
                        writer.WriteEndElement()
                        oRecord2.MoveNext()
                    End While
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord2)
                oRecord2 = Nothing
                GC.Collect()
                writer.WriteEndElement()

                writer.WriteEndElement()
                oRecord.MoveNext()
            End While
        End If

        ''Cierre detalles
        writer.WriteEndElement()
        ''Cierre Factura
        writer.WriteEndElement()
        writer.WriteEndDocument()
        writer.Close()
    End Sub

    Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement(pID)

        writer.WriteString(pName)
        writer.WriteEndElement()

    End Sub

    Private Sub cargarInicial(oCompany As SAPbobsCOM.Company, APP As SAPbouiCOM.Application)
        Try
            If My.Computer.FileSystem.FileExists(Application.StartupPath & "\carga.xlsx") Then
                Dim dataTable As New DataTable
                Dim aValidValues As New ArrayList
                Dim oValid As New validValues
                Dim insertar As Boolean = False
                Dim conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;data source=" & Application.StartupPath & "\" & "carga.xlsx" & ";Extended Properties='Excel 12.0 Xml;HDR=Yes'")
                Dim myDataAdapter As New OleDbDataAdapter("Select * from [" + "Table 1" + "$]", conn)
                myDataAdapter.Fill(dataTable)

                For Each fila As DataRow In dataTable.Rows
                    Dim objeto = fila(14).ToString
                    Dim oValue = fila(0).ToString
                    If objeto = "OWHT" Then
                        Dim oTypeNum = Nothing
                        If fila(8).ToString = "Neto" Then
                            oTypeNum = SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net
                            SetNewTax(fila(1), fila(3).ToString, SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, Double.Parse(fila(9).ToString), fila(10), "99999", fila(13).ToString, IIf(fila(7).ToString = "", 0, Double.Parse(fila(7).ToString)))
                        ElseIf fila(8).ToString = "IVA" Then
                            SetNewTax(fila(1), fila(3).ToString, SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_VAT, Double.Parse(fila(9).ToString), fila(10), "99999", fila(13).ToString, IIf(fila(7).ToString = "", 0, Double.Parse(fila(7).ToString)))
                        End If

                    End If
                Next
                If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Sustentos.txt") = True And My.Computer.FileSystem.FileExists(Application.StartupPath & "\Comprobantes.txt") = True Then
                    updateValidValues()
                End If              
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

End Class
