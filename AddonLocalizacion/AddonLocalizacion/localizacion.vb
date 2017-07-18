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

            'Dim oRecordA As SAPbobsCOM.Recordset
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
            SetFomsUDO()

            UDT_UF.SBOApplication = Me.SBO_Application
            UDT_UF.Company = Me.oCompany
            'Dim facturaNueva As New generarRetencionXML
            'facturaNueva.generaXML("39", "RTNC", oCompany, SBOApplication)
            ' PROBAR()
            Dim oRecord As SAPbobsCOM.Recordset
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("SELECT ISNULL(A.U_STATUS,'N') FROM  [@INF_APP] A")
            If oRecord.RecordCount > 0 Then
                If oRecord.Fields.Item(0).Value <> "I" Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()
                    cargarInicial(oCompany, SBO_Application)
                End If
            Else
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                oRecord = Nothing
                GC.Collect()
                cargarInicial(oCompany, SBO_Application)
            End If
            SBOApplication.StatusBar.SetText("AddOn de LOCALIZACIÓN iniciado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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

            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "pCli"
            oCreationPackage.String = "Retenciones en Venta"
            oMenus.AddEx(oCreationPackage)
          
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "CRtn"
            oCreationPackage.Position = "2"
            oCreationPackage.String = "Generar ATS"
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

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "gui"
            oCreationPackage.Position = "5"
            oCreationPackage.String = "Guía de Remision"
            oMenus.AddEx(oCreationPackage)

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "cpf"
            oCreationPackage.Position = "5"
            oCreationPackage.String = "Cheques Posfechados"
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
            oCreationPackage.String = "Generar ATS"
            oMenus.AddEx(oCreationPackage)

            oMenuItem = SBO_Application.Menus.Item("gui")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GRe"
            oCreationPackage.String = "Guía de Remisión"
            oMenus.AddEx(oCreationPackage)

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "guiM"
            oCreationPackage.String = "Guía de Remisión Masiva"
            oMenus.AddEx(oCreationPackage)

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "tran"
            oCreationPackage.String = "Transportista"
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
            oCreationPackage.UniqueID = "ss"
            oCreationPackage.String = "Series"
            oMenus.AddEx(oCreationPackage)

            oMenuItem = SBO_Application.Menus.Item("cpf")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "cp"
            oCreationPackage.String = "Cheques Posfechados"
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
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
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
        oFilter.AddEx("170")
        oFilter.AddEx("GREMISION")
        oFilter.AddEx("65307")        
        'oFilter.AddEx("139") 'Orders Form
        'oFilter.AddEx("133") 'Invoice Form
        'oFilter.AddEx("169") 'Main Menu
        SBO_Application.SetFilter(oFilters)

    End Sub


    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent

        Try
            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
               
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
                                    SBOApplication.SetStatusBarMessage("El proveedor debe de comenzar con PL o PE", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    System.Media.SystemSounds.Asterisk.Play()
                                    BubbleEvent = False
                                    Return
                                End If
                            End If
                        End If

                        If oTipoIden.Value = "" Then
                            SBOApplication.SetStatusBarMessage("Debe de Elegir un Tipo de Identificación", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            System.Media.SystemSounds.Asterisk.Play()
                            BubbleEvent = False
                            Return
                        End If

                        ''Cuando se selecciona un RUC                      
                        If oTipoIden.Value.Trim = "01" And oBPcode.Value <> "" Then
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

                                'MessageBox.Show(oDocumento.ToString.Chars(2))
                                Dim claserum = Integer.Parse(oDocumento.Value.ToString.Chars(2))
                                If claserum = 9 Then
                                    BubbleEvent = digitoVerificador(oDocumento.Value, Me.SBO_Application, True)
                                Else
                                    If claserum = 6 Then
                                        BubbleEvent = digitoVerificadorPublico(oDocumento.Value, SBOApplication, True)
                                    Else
                                        BubbleEvent = digitoVerificadorIndividual(oDocumento.Value, SBOApplication, True)
                                    End If
                                End If
                            Else
                                SBOApplication.SetStatusBarMessage("RUC debe contener 13 dígitos", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                System.Media.SystemSounds.Asterisk.Play()
                                BubbleEvent = False
                            End If
                        Else
                            If oTipoIden.Value.Trim = "02" Then
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

                                    Dim claserum = Integer.Parse(oDocumento.Value.ToString.Chars(2))
                                    If claserum = 9 Then
                                        BubbleEvent = False
                                        SBOApplication.SetStatusBarMessage("Cédula no válida.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        Return
                                    Else
                                        If claserum = 6 Then
                                            BubbleEvent = digitoVerificadorPublico(oDocumento.Value, SBOApplication, True)
                                        Else
                                            BubbleEvent = digitoVerificadorIndividual(oDocumento.Value, SBOApplication, True)
                                        End If
                                    End If

                                End If
                            Else
                                If oTipoIden.Value.Trim = "03" Then
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
                                Else
                                    If oTipoIden.Value.Trim = "04" Then
                                        Dim oDocumento As SAPbouiCOM.EditText
                                        oDocumento = oUform.Items.Item("U_DOCUMENTO").Specific
                                        oDocumento.Value = Trim(Right(oBPcode.Value, Len(oBPcode.Value) - 2)).ToString
                                        If oDocumento.Value.Trim <> "9999999999999" Then
                                            SBOApplication.SetStatusBarMessage("L(800) Para consumidor final debe de ingresar 9999999999999 ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
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
                            Dim oAplicarRetencion As SAPbouiCOM.ComboBox
                            'Dim oAutoRetencion As SAPbouiCOM.EditText

                            ' oSEstable = oUForm.Items.Item("U_SERIE_ESTABLE").Specific
                            'optoEmision = oUForm.Items.Item("U_PTO_EMISION").Specific
                            'oEstableReten = oUForm.Items.Item("U_STBLE_RETENCION").Specific
                            'optoRetencion = oUForm.Items.Item("U_PTO_RETENCION").Specific
                            oSusTribu = oUForm.Items.Item("U_SUS_TRIBU").Specific
                            oTipoComro = oUForm.Items.Item("U_TI_COMPRO").Specific
                            oNumRetencion = oUForm.Items.Item("U_RETENCION_NO").Specific
                            oAplicarRetencion = oUForm.Items.Item("U_A_APLICARR").Specific

                            If UDT_UF.code <> "" Then
                                oNumRetencion.Value = UDT_UF.code
                            End If

                            ' oAutoRetencion = oUForm.Items.Item("U_AUTORI_RETENCION").Specific
                            ' Else
                            '   SBOApplication.SetStatusBarMessage("Punto de emisión establecimiento debe de tener 3 digitos. ejemp 001", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
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

                'If pVal.FormTypeEx = "133" And pVal.Before_Action = True And pVal.ItemUID = "1" Then
                ' If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                'Try
                'Dim oAutoRetencion As SAPbouiCOM.EditText
                'Dim oUForm = SBOApplication.Forms.GetForm("-133", pVal.FormTypeCount)
                ' oAutoRetencion = oUForm.Items.Item("U_RETENCION_NO").Specific
                'oAutoRetencion.Value = UDT_UF.code
                ' Catch ex As Exception
                'SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                'End Try
                'End If
                '  End If
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

                End If

            End If


            If pVal.FormUID = "GREMISION_" And pVal.ItemUID = "1" And pVal.BeforeAction = True Then

            End If

            If pVal.FormTypeEx = "170" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.BeforeAction = False Then
                Dim oNewItem As SAPbouiCOM.Item
                Dim NewButton As SAPbouiCOM.Button
                Dim oitem As SAPbouiCOM.Item
                Dim ocmdFirma As SAPbouiCOM.Item
                Dim myForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                oitem = myForm.Items.Item("2")
                oNewItem = myForm.Items.Add("btnPago", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                oNewItem.Left = oitem.Left + 80
                oNewItem.Width = oitem.Width + 30
                oNewItem.Top = oitem.Top
                oNewItem.Height = oitem.Height
                NewButton = oNewItem.Specific
                NewButton.Caption = "Retenciones en venta"
                ocmdFirma = myForm.Items.Item("btnPago")
                ocmdFirma.Enabled = True
                BubbleEvent = False
                Return
            End If

            If pVal.FormTypeEx = "170" And pVal.ItemUID = "btnPago" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                Dim opago As New pago_retencion_cliente
                BubbleEvent = False
                Return
            End If

            If pVal.FormUID = "T_GTRANSPORTISTA_" And pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = 2 Then
                Dim myForm = SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                Dim validar As Boolean = False
                Dim oRuc As SAPbouiCOM.EditText
                oRuc = myForm.Items.Item("Item_1").Specific

                Try
                    Double.Parse(oRuc.Value.Trim)
                Catch ex As Exception
                    SBOApplication.SetStatusBarMessage("El número de RUC/CI es incorrecto", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False
                    Return
                End Try

                If oRuc.Value.Trim.Count = 13 Then
                    Dim claserum = Integer.Parse(oRuc.Value.ToString.Chars(2))
                    If claserum = 6 Then
                        If digitoVerificadorPublico(oRuc.Value, Me.SBO_Application, False) = False Then
                            Me.SBO_Application.SetStatusBarMessage("RUC no válido", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                            Return
                        End If
                    Else
                        If claserum = 9 Then
                            If digitoVerificador(oRuc.Value.Trim, Me.SBO_Application, False) = False Then
                                Me.SBO_Application.SetStatusBarMessage("RUC no válido", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                Return
                            End If
                        Else
                            If digitoVerificadorIndividual(oRuc.Value, Me.SBO_Application, False) = False Then
                                Me.SBO_Application.SetStatusBarMessage("RUC no válido", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                Return
                            End If
                        End If
                    End If
                Else
                    If oRuc.Value.Trim.Count = 10 Then
                        Dim claseci = Integer.Parse(oRuc.Value.ToString.Chars(2))
                        If claseci = 6 Then
                            If digitoVerificadorPublico(oRuc.Value, Me.SBO_Application, True) = False Then
                                Me.SBO_Application.SetStatusBarMessage("C.I. no válido", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                Return
                            End If
                        Else
                            If claseci = 9 Then
                                If digitoVerificador(oRuc.Value, Me.SBO_Application, True) = False Then
                                    Me.SBO_Application.SetStatusBarMessage("C.I. no válido", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                    Return
                                End If
                            Else
                                If digitoVerificadorIndividual(oRuc.Value, Me.SBO_Application, True) = False Then
                                    Me.SBO_Application.SetStatusBarMessage("C.I. no válido", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                    Return
                                End If
                            End If
                        End If
                    Else
                        Me.SBO_Application.SetStatusBarMessage("C.I/RUC no válido", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    End If
                End If
                Return
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
                xmlRetencion.generaXML(docEntrynode.InnerText, "RTNC", oCompany, SBOApplication)
                Dim xmlReembolso As New generarFRXML
                xmlReembolso.generarXML(docEntrynode.InnerText, "FR", oCompany, SBOApplication)
            End If
            If pVal.FormTypeEx = "133" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then               
                Dim doc As New XmlDocument
                doc.LoadXml(pVal.ObjectKey)
                Dim docEntrynode = doc.DocumentElement.SelectSingleNode("/DocumentParams/DocEntry")
                Dim tipoFac = tipoFactura(docEntrynode.InnerText)
                Dim oFac As New generarFXML
                oFac.generarXML(docEntrynode.InnerText, "13", oCompany, SBOApplication)
            End If
            If pVal.FormTypeEx = "179" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then               
                Dim doc As New XmlDocument
                doc.LoadXml(pVal.ObjectKey)
                Dim docEntrynode = doc.DocumentElement.SelectSingleNode("/DocumentParams/DocEntry")
                Dim oNC As New generarNCXML
                oNC.generarXML(docEntrynode.InnerText, "14", oCompany, SBOApplication)
            End If

            If pVal.FormTypeEx = "65307" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                Dim doc As New XmlDocument
                doc.LoadXml(pVal.ObjectKey)
                Dim docEntrynode = doc.DocumentElement.SelectSingleNode("/DocumentParams/DocEntry")
                Dim oFacturaExportacion As New generarFEXML
                oFacturaExportacion.generarXML(docEntrynode.InnerText, "13E", oCompany, SBOApplication)
            End If

            If pVal.FormTypeEx = "65303" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                Dim doc As New XmlDocument
                doc.LoadXml(pVal.ObjectKey)
                Dim docEntrynode = doc.DocumentElement.SelectSingleNode("/DocumentParams/DocEntry")
                Dim oNotaDebito As New generarNDXML
                oNotaDebito.generarXML(docEntrynode.InnerText, "ND", oCompany, SBO_Application)
            End If
            If pVal.FormTypeEx = "181" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                Dim doc As New XmlDocument
                doc.LoadXml(pVal.ObjectKey)
                Dim docEntrynode = doc.DocumentElement.SelectSingleNode("/DocumentParams/DocEntry")
                Dim oRecord As SAPbobsCOM.Recordset
                oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord.DoQuery("exec ACTUALIZANOTACREDITOPROVEEDORES " & docEntrynode.InnerText & "")
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                oRecord = Nothing
                GC.Collect()
            End If
            If pVal.FormTypeEx = "65306" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                Dim doc As New XmlDocument
                doc.LoadXml(pVal.ObjectKey)
                Dim docEntrynode = doc.DocumentElement.SelectSingleNode("/DocumentParams/DocEntry")
                Dim oRecord As SAPbobsCOM.Recordset
                oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord.DoQuery("exec ACTUALIZANOTADEBITOPROVEEDORES " & docEntrynode.InnerText & "")
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                oRecord = Nothing
                GC.Collect()
            End If
            If pVal.FormTypeEx = "GREMISION" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                              
                Dim oRecord As SAPbobsCOM.Recordset
                Dim docEntry As String = ""
                oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord.DoQuery("exec ACTUALIZAR_DOC_GUIA")

                If oRecord.RecordCount > 0 Then
                    docEntry = oRecord.Fields.Item(0).Value
                End If
                Dim generaXml As New generarGRXML
                generaXml.generarXML(docEntry, "GR", oCompany, SBOApplication)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                oRecord = Nothing
                GC.Collect()
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
        If (pVal.MenuUID = "cp") And (pVal.BeforeAction = False) Then
            Dim cheques As New cheques_posfechados
            BubbleEvent = False
        End If
        If (pVal.MenuUID = "pCli") And (pVal.BeforeAction = False) Then
            Dim pagos_Remision As New pago_retencion_cliente
            BubbleEvent = False
        End If
        If (pVal.MenuUID = "GRe") And (pVal.BeforeAction = False) Then
            Dim Guia As New guia_remision
            BubbleEvent = False
        End If

        If (pVal.MenuUID = "guiM") And (pVal.BeforeAction = False) Then
            Dim GuiaMasiva As New guia_remision_masiva
            BubbleEvent = False
        End If
        If (pVal.MenuUID = "infCl") And (pVal.BeforeAction = False) Then
            Dim infoCliente As New retencion_cliente
            BubbleEvent = False
        End If
        If (pVal.MenuUID = "tran") And (pVal.BeforeAction = False) Then
            Dim trans As New transportista
            BubbleEvent = False
        End If



        If (pVal.MenuUID = "infTri") And (pVal.BeforeAction = False) Then
            Dim inf As New inf_tributaria
            BubbleEvent = False
        End If
        If (pVal.MenuUID = "rete") And (pVal.BeforeAction = False) Then
            Dim oATS As New generar_ATS
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
            Return
        End If
        If (pVal.MenuUID = "rinv") And (pVal.BeforeAction = False) Then
            Dim ventas As New retencion_info_ventas
            BubbleEvent = False
        End If
    End Sub

    Private Sub SetNewItems()
        Try

            UDT_UF.userField(oCompany, "OCRD", "Tipo de Identificación", 45, "IDENTIFICACION", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBO_Application)
            'UDT_UF.userField(oCompany, "OCRD", "TIPO RUC", 45, "TIPO_RUC", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "Parte Relacionada", 25, "PT_RELACIO", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "Tipo Cliente/Proveedor", 45, "T_C_P", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "Pago a Residente o No Residente  ", 45, "TIPO_CONTRI", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "Tipo de Régimen Fiscal del Exterior", 25, "T_R_FISCAL", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "País al que se le Realiza el Pago Régimen General", 25, "P_R_FISCAL", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "País al que se le Realiza el Pago Paraíso Fiscal", 25, "P_R_PFISCAL", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "Denominación del Régimen Fiscal", 30, "D_R_FISCAL", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "País al que Efectúa el Pago", 45, "PAIS_PAGO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "Aplica Conveninio de Doble Tributación", 25, "D_TRIBUTACION", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "No. Documento", 45, "DOCUMENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "OCRD", "Pago Ext. Sujeto a Retención", 25, "SUJE_RETENCION", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)



            ' UDT_UF.userField(oCompany, "OCRD", "RISE", 45, "RISE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "Clase de Sujeto", 45, "TIPO_SUJETO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "Origen de Ingresos", 3, "O_INGRESOS", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "Estado Civil", 3, "ESTADO_CIVIL", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OCRD", "Sexo", 3, "SEXO", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)



            'UDT_UF.userField(oCompany, "OCRD", "TIPO SUJETO", "TIPO_SUJETO", )
            UDT_UF.userTable(oCompany, "INF_TRIBUTARIA", "INFORMACION TRIBUTARIA", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "INF_PARTNER", "ADICIONAL AL PARTNER", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_PARTNER", "CLIENTE", 15, "CLIENTE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_PARTNER", "RAZON", 70, "RAZON", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_PARTNER", "BASE", 10, "BASE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_PARTNER", "RETENCION", 10, "RETENCION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_PARTNER", "CUENTA BASE", 10, "B_CUENTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_PARTNER", "CUENTA RETENCION", 10, "R_CUENTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)


            UDT_UF.userTable(oCompany, "INF_APP", "INFORMACION DE LA APLICACION", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_APP", "VERSION", 12, "VERSION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_APP", "FECHA", 12, "FECHA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_APP", "STATUS", 12, "STATUS", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
           

            UDT_UF.userTable(oCompany, "MUNI_CANTO", "CANTON O MUNICIPIO", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "PARROQUIAS", "PARROQUIAS", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userField(oCompany, "@PARROQUIAS", "CANTON", 30, "CANTON", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@PARROQUIAS", "PRIVINCIA", 30, "PROVINCIA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            'UDT_UF.userField(oCompany, "@INF_PARTNER", "DOBLE TRIBU", 30, "DO_TRI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
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


            ''Guia de Remision Encabezado
            UDT_UF.userTable(oCompany, "GREMISION", "GUIA DE REMISION", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_Document, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "RUC DESTINATARIO", 35, "RUC_DESTI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "DESTINATARIO", 30, "RUC_DESTI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "No. DOCUMENTO", 30, "N_DOCUMENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "PUNTO DE LLEGADA", 30, "PTO_LLEGADA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "PUNTO DE PARTIDA", 30, "PTO_PARTIDA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "FECHA INICIO TRASLADO", 30, "F_ITRASLADO", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "FECHA FIN TRASLADO", 30, "F_FTRASLADO", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "GUIA TRASPORTISTA FINAL", 30, "G_TFINAL", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "CONSECUTIVO", 30, "CONSECUTIVO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "RUC TRANSPORTISTA", 30, "G_TRANSPOR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "TRANSPORTISTA", 60, "TRANPORTISTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "PLACA", 30, "PLACA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "BIENES TRANSPORTADOS", 30, "B_TRANS", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "FECHA ENVIO", 30, "F_ENVIO", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION", "RUTA", 30, "G_RUTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            'Guia de remision detalle
            UDT_UF.userTable(oCompany, "DGREMISION", "DETALLE GUIA DE REMISION", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_DocumentLines, False, SBOApplication)
            UDT_UF.userField(oCompany, "@DGREMISION", "TIPO DE DOCUMENTO", 6, "TIPO_DOC", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@DGREMISION", "DOC. INICIAL", 6, "DOC_INI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@DGREMISION", "DOC. FINAL", 6, "FINAL", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@DGREMISION", "NO. EMPAQUE", 6, "N_EMPAQUE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            'Cheques posfechados Encabezado.
            UDT_UF.userTable(oCompany, "CPEPOSFE", "CHEQUE POSFECHADO ENCABEZADO", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_Document, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPEPOSFE", "RUC", 25, "RUC", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPEPOSFE", "NOMBRE", 50, "NOMBRE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPEPOSFE", "NO.FACTURA", 10, "N_FACTURA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPEPOSFE", "INSTITUCION FINANCIERA", 55, "I_FINANCIERA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPEPOSFE", "NUMERO DE CHEQUE", 55, "N_CHEQUE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPEPOSFE", "MONTO", 25, "MONTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPEPOSFE", "EMISION CHEQUE", 25, "E_CHEQUE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPEPOSFE", "FECHA DE COBRO", 25, "FC_CHEQUE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPEPOSFE", "FECHA EFECTIVA", 25, "FE_CHEQUE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPEPOSFE", "ASESOR", 55, "ASESOR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            'Cheques posfechados Detalle
            UDT_UF.userTable(oCompany, "CPDPOSFE", "CHEQUE POSFECHADO DETALLE", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_DocumentLines, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPDPOSFE", "NUMERO FACTURA", 25, "N_FACTURA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPDPOSFE", "FECHA FACTURA", 25, "F_FACTURA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPDPOSFE", "MONTO FACTURA", 25, "M_FACTURA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPDPOSFE", "VALOR COBRADO", 25, "V_COBRADO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPDPOSFE", "CUOTA", 25, "CUOTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@CPDPOSFE", "COMENTARIOS", 25, "COMENTARIOS", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)


            UDT_UF.userTable(oCompany, "TGUIA", "TIPO GUIA", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines, False, SBOApplication)
            UDT_UF.userField(oCompany, "@TGUIA", "DOCUMENTO INICIAL", 14, "DOC_INI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@TGUIA", "DOCUMENTO FINAL", 14, "DOC_FI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@TGUIA", "NO. EMPAQUE", 14, "N_EMPAQUE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userTable(oCompany, "T_GTRANSPORTISTA", "TRANSPORTISTAS", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_MasterData, False, SBOApplication)
            UDT_UF.userField(oCompany, "@T_GTRANSPORTISTA", "PLACA", 35, "PLACA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userTable(oCompany, "T_GTRANSPLACA", "PLACA TRANSPORTISTA", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines, False, SBOApplication)
            UDT_UF.userField(oCompany, "@T_GTRANSPLACA", "PLACA", 45, "PLACA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)


            UDT_UF.userTable(oCompany, "P_RETENCION", "PAGO RETENCION", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement, False, SBOApplication)
            UDT_UF.userField(oCompany, "@P_RETENCION", "RUC CLIENTE", 20, "CLIENTE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@P_RETENCION", "NO DOCUMENTO", 20, "N_DOCUMENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@P_RETENCION", "BASE", 20, "BASE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@P_RETENCION", "RETENCION", 20, "RETENCION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@P_RETENCION", "TOTAL BASE", 20, "T_BASE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@P_RETENCION", "RETENCION", 20, "T_RETENCION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@P_RETENCION", "CUENTAB", 12, "CUENTAB", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@P_RETENCION", "CUENTAR", 12, "CUENTAR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@P_RETENCION", "ESTADO", 3, "ESTADO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userTable(oCompany, "INFP_RETENCION", "INFO. PAGO RETENCION", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INFP_RETENCION", "AUTORIZACION", 60, "AUTORIZACION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INFP_RETENCION", "NUMEROA", 13, "NUMEROA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INFP_RETENCION", "FECHAR", 10, "FECHAR", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INFP_RETENCION", "FECHAC", 10, "FECHAC", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INFP_RETENCION", "CLIENTE", 30, "CLIENTE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INFP_RETENCION", "DOCUMENTO", 10, "DOCUMENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)


            UDT_UF.userTable(oCompany, "SERIES", "INFORMACION AUTORIZACIONES", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "SERIE", 3, "SERIE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "NO AUTORIZACION", 45, "NO_AUTORI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "FECHA CADUCIDAD", 45, "FECHA_CADU", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "DIRECION", 70, "DIRECCION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "CIUDAD", 70, "CIUDAD", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "TELEFONO", 70, "TELEFONO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@SERIES", "DIGITAL", 10, "DIGITAL", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)




            UDT_UF.userField(oCompany, "OWHT", "CODIGO ATS", 45, "COD_ATS", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userField(oCompany, "OPCH", "Sustento Tributario", 60, "SUS_TRIBU", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Tipo de Comprobante", 45, "TI_COMPRO", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OINV", "Forma de Pago", 60, "FORMA_PAGO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OINV", "Número de Autorización", 60, "NO_AUTORI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Número de Retención", 45, "RETENCION_NO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Número de autorizacion Retencion", 60, "NA_RETENCION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Fecha de Retención", 60, "F_RETENCION", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Aplicar Retención", 6, "A_APLICARR", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            ' Notas de crédito y de debtito
            UDT_UF.userField(oCompany, "OPCH", "Doc. Modificado", 45, "D_MODIFICADO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Fecha Doc. Modificado", 45, "F_MODIFICADO", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Tipo Comprobante Doc. Modificado", 5, "T_COMPROBANTEM", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Número Autorización Doc. Modificado", 60, "N_AUTORIZACIONM", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            'Factura de Exportación
            UDT_UF.userField(oCompany, "OPCH", "Tipo de Exportación", 3, "T_EXPORT", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Tipo de Ingresos del Exterior", 25, "T_INGRE_EXT", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Impuesto a la Renta o Similar Ext", 25, "IMPUESTO_RENTA", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Valor IR o Similar en el Ext", 25, "V_IR_SIMI", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Fecha Embarque", 25, "F_EMBARQUE", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Valor FOB", 25, "V_FOB", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Número de Transporte Ext.", 13, "N_TRASNPORT", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Distrito Aduanero", 13, "D_ADUANERO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Año", 13, "ANO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Regimen", 13, "REGIMEN", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Correlativo", 10, "CORRELATIVO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)


            UDT_UF.userField(oCompany, "OPCH", "Incoterm Factura", 45, "INCO_TERM", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Lugar Incoterm", 45, "LUGAR_INCOTERM", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "País Origen", 45, "PAIS_ORIGEN", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Puerto Embarque", 45, "PUERTO_EMBARGUE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Puerto Destino", 45, "PUERTO_DESTINO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Pais Destino", 45, "PAIS_DESTINO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "País Adquisición", 45, "PAIS_ADQUISION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "IncoTerm Total SinImpuestos", 45, "TERM_TOT_SIN_IMPUESTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Flete Internacional", 45, "FLETE_INTERNA", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Seguro Internacional", 45, "SEGURO_INTERNA", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Gastos Aduaneros", 45, "GASTOS_ADUANEROS", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Gastos Transporte", 45, "G_TRANS_OTROS", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Guia Remision", 45, "G_REMISION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Transportista", 60, "TRANSPORTISTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)


            UDT_UF.userField(oCompany, "OPCH", "Fecha pago Dividendo", 60, "P_DIVIDENDO", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Año Dividendo", 6, "A_DIVIDENDO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "IR Asoc. Dividendo", 4, "IR_DIVIDENDO", SAPbobsCOM.BoFieldTypes.db_Numeric, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Cant. Caja Banano", 4, "C_BANANO", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "OPCH", "Precio Caja Banano", 4, "PC_BANANO", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)

            'CAMPOS PARA FACTURA DE REEEMBOLSO
            UDT_UF.userField(oCompany, "PCH1", "Tipo de Identificación  ", 13, "T_IDENTIFICACION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Número de Identificación", 13, "N_RUC", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Nombre de Proveedor ", 60, "N_PROVEEDOR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Tipo de Comprobante", 11, "T_COMPROBANTE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Serie Comprobante", 3, "SE_ESTABLE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Punto de Emision", 3, "PTO_EMISION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Número de Comprobante", 11, "N_FACTURA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Fecha Comprobante", 11, "FE_EMISION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "N° Autorización de Comprobante", 60, "AUTO_REEMBOLSO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Base imponible tarifa 0% de IVA", 6, "BASE_0", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Base imponible tarifa IVA diferente de 0% ", 6, "BASE_12", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Base imponible no objeto de IVA", 6, "IVA_NOBJETOREEM", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Base imponible exenta de IVA ", 6, "B_EXENTA_REEM", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Monto IVA ", 11, "MONTO_IVA", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Monto ICE ", 11, "MONTO_ICE", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Motivo", 60, "MOTIVO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "Guia de Proveedor", 11, "GUIA_PROVEEDOR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "TIPO PROVEEDOR ", 13, "T_PROVEEDOR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "PAIS PROVEEDOR", 13, "PA_PROVEEDOR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "PCH1", "ID PROVEEDOR REEMBOLSO ", 13, "ID_PROVEEDOR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OACT", "RUC de Banco  ", 13, "RUC_BANCO ", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OACT", "Cuenta  de Retención", 13, "CUENTA_RET ", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OJDT", "Aplica para ATS ", 13, "APLICA_ATS", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBOApplication)
            UDT_UF.userField(oCompany, "OITM", "ICE por Unidad", 13, "ICE_UNIDAD", SAPbobsCOM.BoFieldTypes.db_Float, False, SBOApplication)
            UDT_UF.userField(oCompany, "OITM", "Codigo ICE", 6, "ICE_CODIGO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            'GUIA DE REMISION MASIVA ENCABEZADO
            UDT_UF.userTable(oCompany, "GREMISION_M", "GUIA DE REMISION MASIVA", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_Document, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "SERIE", 19, "SERIE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "PUNTO PARTIDA", 30, "P_PARTIDA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "TRANSPORTISTA", 30, "R_TRANSPORTISTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "RAZON TRANSPORTISTA", 30, "RA_TRANSPORTISTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "CLIENTE", 30, "CLIENTE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "PLACA", 30, "PLACA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "FECHA INICIAL", 30, "F_INICIAL", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "FECHA FINAL", 30, "F_FINAL", SAPbobsCOM.BoFieldTypes.db_Date, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "TIPO DOCUMENTO", 30, "T_DOCUMENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "MOTIVO TRASLADO", 60, "MOTIVO_TRASLADO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "FACTURA DESDE", 60, "FACDE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@GREMISION_M", "FACTURA HASTA", 60, "FACHASTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userTable(oCompany, "G_ULTIMO", "PAGO RETENCION", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement, False, SBOApplication)
            UDT_UF.userField(oCompany, "@G_ULTIMO", "ID GUIA", 13, "DOCENTRY", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)


            'ATS TABLES
            UDT_UF.userTable(oCompany, "LC_TIPO_IDENTIFI", "TIPO IDENTIFICACION", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "LC_TIPO_TRANSACTION", "TIPO TRANSACCION", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "LC_PERIODO", "PERIODO", 45, "PERIODO", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "LC_DISTRI_ADUANERO", "DISTRITO ADUANERO", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "LC_COD_REGIMEN", "CODIGO REGIMEN", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "LC_TARJ_CREDITO", "TARJETA DE CREDITO", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "PAIS", "REGISTRO DE PAIS", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "LC_IMPUESTO", "CODIGO IMPUESTO", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userTable(oCompany, "LC_PORCENTAJE", "PORCENTAJE DE IMPUESTO", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)

            UDT_UF.userTable(oCompany, "LC_T_FIDEICOMISOS", "TIPO FIDEICOMISOS", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userField(oCompany, "@LC_T_FIDEICOMISOS", "CODIGO PORCENTAJE", 5, "PORCENTAJE", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@LC_T_FIDEICOMISOS", "ACTIVO", 3, "ACTIVO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userTable(oCompany, "LC_T_EXPOR", "TIPO EXPORTACION/INGRESO", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userField(oCompany, "@LC_T_EXPOR", "ACTIVO", 3, "ACTIVO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userTable(oCompany, "LC_P_RETEN_IVA", "PORCENTAJE RETENCION IVA", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userField(oCompany, "@LC_P_RETEN_IVA", "ACTIVO", 3, "ACTIVO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)

            UDT_UF.userTable(oCompany, "LC_PARAISO_FISCAL", "PARAISO FISCAL", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userField(oCompany, "@LC_PARAISO_FISCAL", "PAIS LIGADO", 100, "P_LIGADO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@LC_PARAISO_FISCAL", "CODIGO PAIS", 4, "C_PAIS", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "@LC_PARAISO_FISCAL", "ACTIVO", 3, "ACTIVO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)


            'UDT_UF.userField(oCompany, "OPCH", "SECUENCIAL", 3, "PTO_EMISION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            ''updateValidValues()
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
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\Sustentos.txt")
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

            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\Comprobantes.txt")
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

            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\ComprobantesModi.txt")
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


            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\Cantones.txt")
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


            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\PAIS.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()

                    Dim oRecord As SAPbobsCOM.Recordset
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim SQL = "INSERT INTO [@PAIS] VALUES ('" & currentLine(1) & "','" & currentLine(0).ToString & "')"
                    oRecord.DoQuery(SQL)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()

                End While
            End Using
            validArrayList.Clear()
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\Formas_Pago.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()
                    Dim oValidV As New validValues
                    oValidV.value = currentLine(1).ToString
                    oValidV.descrip = currentLine(0).ToString
                    validArrayList.Add(oValidV)
                End While
                UDT_UF.updateUserField(oCompany, "OINV", "FORMA_PAGO", validArrayList)
                validArrayList.Clear()
            End Using


            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\TIPO_TRANSACCION.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()
                    Dim oRecord As SAPbobsCOM.Recordset
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim SQL = "INSERT INTO [@LC_TIPO_TRANSACTION] VALUES ('" & currentLine(0) & "','" & currentLine(1).ToString & "')"
                    oRecord.DoQuery(SQL)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()
                End While
            End Using

            '-------------------
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\TIPO_IDENTIFICACION.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()
                    Dim oRecord As SAPbobsCOM.Recordset
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim SQL = "INSERT INTO [@LC_TIPO_IDENTIFI] VALUES ('" & currentLine(0) & "','" & currentLine(1).ToString & "')"
                    oRecord.DoQuery(SQL)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()
                End While
            End Using

            '--------------
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\DISTRITO_ADUANERO.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()
                    Dim oRecord As SAPbobsCOM.Recordset
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim SQL = "INSERT INTO [@LC_DISTRI_ADUANERO] VALUES ('" & currentLine(0) & "','" & currentLine(1).ToString & "')"
                    oRecord.DoQuery(SQL)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()
                End While
            End Using

            '---------------------

            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\PERIODO.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()
                    Dim oRecord As SAPbobsCOM.Recordset
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim SQL = "INSERT INTO [@LC_PERIODO] VALUES ('" & currentLine(1) & "','" & currentLine(0).ToString & "')"
                    oRecord.DoQuery(SQL)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()
                End While
            End Using


            '--------------------------------------------
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\CODIGO_REGIMEN.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()
                    Dim oRecord As SAPbobsCOM.Recordset
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim SQL = "INSERT INTO [@LC_COD_REGIMEN] VALUES ('" & currentLine(0) & "','" & currentLine(1).ToString & "')"
                    oRecord.DoQuery(SQL)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()
                End While
            End Using

            '--------------------------------------------
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\TARJETA_CREDITO.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                While Not fileReader.EndOfData
                    Dim currentLine As String() = fileReader.ReadFields()
                    Dim oRecord As SAPbobsCOM.Recordset
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim SQL = "INSERT INTO [@LC_TARJ_CREDITO] VALUES ('" & currentLine(0) & "','" & currentLine(1).ToString & "')"
                    oRecord.DoQuery(SQL)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()
                End While
            End Using

            '--------------------------------------------
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\PARAISO_FISCAL.txt")
                fileReader.TextFieldType = FileIO.FieldType.Delimited
                fileReader.SetDelimiters(vbTab)
                Dim x As Integer = 0
                Try
                    While Not fileReader.EndOfData
                        Dim currentLine As String() = fileReader.ReadFields()
                        Dim oRecord As SAPbobsCOM.Recordset
                        x = x + 1
                        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim SQL = "INSERT INTO [@LC_PARAISO_FISCAL] VALUES ('" & currentLine(1) & "','" & currentLine(0).ToString & "','" & currentLine(2) & "','" & currentLine(3) & "','" & currentLine(4) & "')"
                        oRecord.DoQuery(SQL)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                        oRecord = Nothing
                        GC.Collect()
                    End While
                Catch ex As Exception
                    MessageBox.Show("linea  " & x)
                End Try
               
            End Using


            '--------------------------------------------
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\FIDEICOMISOS.txt")
                Dim xe As Integer = 0
                Try
                    fileReader.TextFieldType = FileIO.FieldType.Delimited
                    fileReader.SetDelimiters(vbTab)

                    While Not fileReader.EndOfData
                        Dim currentLine As String() = fileReader.ReadFields()
                        Dim oRecord As SAPbobsCOM.Recordset
                        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        xe = xe + 1
                        Dim SQL = "INSERT INTO [@LC_T_FIDEICOMISOS] VALUES ('" & currentLine(0) & "','" & currentLine(1).ToString & currentLine(0) & "','" & currentLine(2) & "','" & currentLine(3) & "')"
                        oRecord.DoQuery(SQL)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                        oRecord = Nothing
                        GC.Collect()
                    End While
                Catch ex As Exception
                    MessageBox.Show("error " & xe & "   " & ex.Message)
                End Try               
            End Using
            '--------------------------------------------
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\IMPUESTO.txt")
                Dim xe As Integer = 0
                Try
                    fileReader.TextFieldType = FileIO.FieldType.Delimited
                    fileReader.SetDelimiters(vbTab)

                    While Not fileReader.EndOfData
                        Dim currentLine As String() = fileReader.ReadFields()
                        Dim oRecord As SAPbobsCOM.Recordset
                        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        xe = xe + 1
                        Dim SQL = "INSERT INTO [@LC_IMPUESTO] VALUES ('" & currentLine(0) & "','" & currentLine(1).ToString & "')"
                        oRecord.DoQuery(SQL)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                        oRecord = Nothing
                        GC.Collect()
                    End While
                Catch ex As Exception
                    MessageBox.Show("error " & xe & "   " & ex.Message)
                End Try
            End Using


            '--------------------------------------------
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\ATS\PORCENTAJE_IMPUESTO.txt")
                Dim xe As Integer = 0
                Try
                    fileReader.TextFieldType = FileIO.FieldType.Delimited
                    fileReader.SetDelimiters(vbTab)

                    While Not fileReader.EndOfData
                        Dim currentLine As String() = fileReader.ReadFields()
                        Dim oRecord As SAPbobsCOM.Recordset
                        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        xe = xe + 1
                        Dim SQL = "INSERT INTO [@LC_PORCENTAJE] VALUES ('" & currentLine(0) & "','" & currentLine(1).ToString & "')"
                        oRecord.DoQuery(SQL)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                        oRecord = Nothing
                        GC.Collect()
                    End While
                Catch ex As Exception
                    MessageBox.Show("error " & xe & "   " & ex.Message)
                End Try
            End Using

            Dim validArray As New ArrayList()
            Dim oValid As New validValues
            oValid.value = "01"
            oValid.descrip = "RUC"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "02"
            oValid.descrip = "CEDULA"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "03"
            oValid.descrip = "PASAPORTE"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "04"
            oValid.descrip = "CONSUMIDOR FINAL"
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
            oValid.descrip = "PAGO A RESIDENTE"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "02"
            oValid.descrip = "PAGO A NO RESIDENTE"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "TIPO_CONTRI", validArray)

            validArray.Clear()

            oValid = Nothing
            oValid = New validValues

            oValid.value = "N"
            oValid.descrip = "NATURAL"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues
            oValid.value = "J"
            oValid.descrip = "JURIDICA"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "TIPO_SUJETO", validArray)

            validArray.Clear()
            oValid = Nothing
            oValid = New validValues
            oValid.value = "SI"
            oValid.descrip = "SI"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues
            oValid.value = "NO"
            oValid.descrip = "NO"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "OBLI_CONTA", validArray)

            validArray.Clear()
            oValid = Nothing
            oValid = New validValues
            oValid.value = "SI"
            oValid.descrip = "SI"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "NO"
            oValid.descrip = "NO"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "PT_RELACIO", validArray)
            validArray.Clear()

            oValid = Nothing
            oValid = New validValues
            oValid.value = "SI"
            oValid.descrip = "SI"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "NO"
            oValid.descrip = "NO"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "D_TRIBUTACION", validArray)


            validArray.Clear()

            oValid = Nothing
            oValid = New validValues
            oValid.value = "M"
            oValid.descrip = "MASCULINO"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "F"
            oValid.descrip = "FEMENINO"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "SEXO", validArray)

            validArray.Clear()

            oValid = Nothing
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "PESONAL NATURAL"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "02"
            oValid.descrip = "SOCIEDAD"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "T_C_P", validArray)




            validArray.Clear()


            oValid = Nothing
            oValid = New validValues
            oValid.value = "SI"
            oValid.descrip = "SI"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "NO"
            oValid.descrip = "NO"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "SUJE_RETENCION", validArray)


            validArray.Clear()


            oValid = Nothing
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "REGIMEN GENERAL"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "02"
            oValid.descrip = "PARAISO FISCAL"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "03"
            oValid.descrip = "PREFERENTE O JURISDICCION DE MENOR IMPOSICION"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "T_R_FISCAL", validArray)

            validArray.Clear()


            oValid = Nothing
            oValid = New validValues
            oValid.value = "B"
            oValid.descrip = "EMPLEADO PUBLICO"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "V"
            oValid.descrip = "EMPLEADO PRIVADO"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "I"
            oValid.descrip = "INDEPENDIENTE"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues
            oValid.value = "A"
            oValid.descrip = "AMA DE CASA O ESTUDIANTE"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues
            oValid.value = "R"
            oValid.descrip = "RENTISTA"
            validArray.Add(oValid)


            oValid = Nothing
            oValid = New validValues
            oValid.value = "H"
            oValid.descrip = "JUBILADO"
            validArray.Add(oValid)


            oValid = Nothing
            oValid = New validValues
            oValid.value = "M"
            oValid.descrip = "REMESAS DEL EXTERIOR"
            validArray.Add(oValid)

            UDT_UF.updateUserField(oCompany, "OCRD", "O_INGRESOS", validArray)

            validArray.Clear()


            oValid = Nothing
            oValid = New validValues
            oValid.value = "S"
            oValid.descrip = "SOLTERO"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "C"
            oValid.descrip = "CASADO"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "D"
            oValid.descrip = "DIVORCIADO"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "U"
            oValid.descrip = "UNION LIBRE"
            validArray.Add(oValid)
            oValid = Nothing
            oValid = New validValues
            oValid.value = "V"
            oValid.descrip = "VIUDO"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OCRD", "ESTADO_CIVIL", validArray)


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
            UDT_UF.updateUserField(oCompany, "OPCH", "A_APLICARR", validArray)
            validArray.Clear()


            oValid = Nothing
            oValid = New validValues
            oValid.value = "01"
            oValid.descrip = "EXPORTACION DE BINES CON REFRENDO"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "02"
            oValid.descrip = "EXPORTACION DE BIENES IN REFRENDO"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues
            oValid.value = "03"
            oValid.descrip = "EXPORTACION DE SERVICIOS U OTROS"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OPCH", "T_EXPORT", validArray)

            validArray.Clear()
            oValid = Nothing
            oValid = New validValues
            oValid.value = "SI"
            oValid.descrip = "SI"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "NO"
            oValid.descrip = "NO"
            validArray.Add(oValid)
            UDT_UF.updateUserField(oCompany, "OPCH", "IMPUESTO_RENTA", validArray)
            validArray.Clear()




            validArray.Clear()

            oValid = Nothing
            oValid = New validValues
            oValid.value = "No Aplica"
            oValid.descrip = "No Aplica"
            validArray.Add(oValid)

            oValid = Nothing
            oValid = New validValues

            oValid.value = "RENTA"
            oValid.descrip = "RENTA"
            validArray.Add(oValid)


            oValid = Nothing
            oValid = New validValues

            oValid.value = "IVA"
            oValid.descrip = "IVA"
            validArray.Add(oValid)

            UDT_UF.updateUserField(oCompany, "OACT", "CUENTA_RET", validArray)



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

            UDT_UF.updateUserField(oCompany, "OJDT", "APLICA_ATS", validArray)

            Dim oRecordA As SAPbobsCOM.Recordset


            oRecordA = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordA.DoQuery("INSERT INTO [@INF_APP] VALUES (1,'LOCALIZACION',1.0,'" & Date.Now.ToString("yyy/MM/dd") & "','I')")
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordA)
            oRecordA = Nothing
            GC.Collect()
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

        oRecord.DoQuery("exec ENCABEZADO_FACTURA '" & DocEntry & "','13'")
        Dim writer As New XmlTextWriter("Comprobante (F) No." & DocEntry.ToString & ".xml", System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("factura")
        writer.WriteAttributeString("id", "comprobante")
        writer.WriteAttributeString("version", "2.0.0")
        writer.WriteStartElement("infoTributaria")
        createNode("razonSocial", oRecord.Fields.Item(2).Value.ToString, writer)
        'createNode("ambiente", oRecord.Fields.Item(0).Value.ToString, writer)
        'createNode("tipoEmision", oRecord.Fields.Item(1).Value.ToString, writer)
        createNode("ruc", oRecord.Fields.Item(3).Value.ToString.PadLeft(13, "0"), writer)
        'createNode("claveAcesso", claveAcceso(oRecord).PadLeft(49, "0"), writer)
        'createNode("claveAcesso", "", writer)
        createNode("codDoc", oRecord.Fields.Item("codDoc").Value.ToString.PadLeft(2, "0"), writer)
        createNode("estab", oRecord.Fields.Item("estable").Value.ToString.PadLeft(3, "0"), writer)
        createNode("ptoEmi", oRecord.Fields.Item("ptoemi").Value.ToString.PadLeft(3, "0"), writer)
        createNode("secuencial", oRecord.Fields.Item("secuencial").Value.ToString.PadLeft(9, "0"), writer)
        createNode("dirMatriz", oRecord.Fields.Item("dirMatriz").Value.ToString, writer)
        Dim direccion = oRecord.Fields.Item("dirMatriz").Value.ToString
        ''Cierre info Tributaria
        writer.WriteEndElement()

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
        oRecord = Nothing
        GC.Collect()

        writer.WriteStartElement("infoFactura")
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecord.DoQuery("exec SP_INFO_FACTURA '" & DocEntry & "','13'")
        createNode("fechaEmision", Date.Parse(oRecord.Fields.Item("DATE").Value.ToString).ToString("dd/MM/yyyy"), writer)
        createNode("dirEstablecimiento", oRecord.Fields.Item(1).Value, writer)
        createNode("contribuyenteEspecial", oRecord.Fields.Item(2).Value, writer)
        createNode("obligadoContabilidad", oRecord.Fields.Item(3).Value, writer)
        createNode("tipoIdentificacionComprador", oRecord.Fields.Item("U_IDENTIFICACION").Value.ToString, writer)
        createNode("guiaRemision", "", writer)
        createNode("razonSocialComprador", oRecord.Fields.Item("CardName").Value.ToString, writer)
        createNode("identificacionComprador", oRecord.Fields.Item("U_DOCUMENTO").Value.ToString, writer)
        createNode("totalSinImpuestos", oRecord.Fields.Item("sin_impuesto").Value.ToString, writer)
        createNode("totalDescuento", oRecord.Fields.Item("totDescuento").Value.ToString, writer)

        writer.WriteStartElement("totalConImpuestos")
        Dim importeTotal = oRecord.Fields.Item("DocTotal").Value.ToString
        Dim moneda = oRecord.Fields.Item("MONEDA").Value.ToString
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
        oRecord = Nothing
        GC.Collect()

        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecord.DoQuery("exec SP_Total_Con_Impuesto '" & DocEntry & "','13'")
        If oRecord.RecordCount > 0 Then
            While oRecord.EoF = False
                writer.WriteStartElement("totalImpuesto")
                createNode("codigo", oRecord.Fields.Item(0).Value.ToString, writer)
                createNode("codigoPorcentaje", oRecord.Fields.Item(1).Value.ToString, writer)
                createNode("baseImponible", oRecord.Fields.Item(2).Value.ToString, writer)
                createNode("tarifa", oRecord.Fields.Item(3).Value, writer)
                createNode("valor", oRecord.Fields.Item(4).Value.ToString, writer)
                writer.WriteEndElement()
                oRecord.MoveNext()
            End While
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
        oRecord = Nothing
        GC.Collect()

        ''Cierre TotalConImpuestos
        writer.WriteEndElement()
        createNode("propina", "0.00", writer)
        createNode("importeTotal", importeTotal, writer)
        createNode("moneda", moneda, writer)

        writer.WriteStartElement("pagos")
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecord.DoQuery("exec SP_Forma_Pago '" & DocEntry & "'")
        If oRecord.RecordCount > 0 Then
            While oRecord.EoF = False
                writer.WriteStartElement("pago")
                createNode("formaPago", oRecord.Fields.Item(0).Value, writer)
                createNode("total", oRecord.Fields.Item(1).Value, writer)
                createNode("plazo", oRecord.Fields.Item(2).Value, writer)
                createNode("unidadTiempo", oRecord.Fields.Item(3).Value, writer)
                writer.WriteEndElement()
                oRecord.MoveNext()
            End While
        End If
        ''Cierre Pagos
        writer.WriteEndElement()



        ''Cierre INFO FACTURA
        writer.WriteEndElement()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
        oRecord = Nothing
        GC.Collect()

        writer.WriteStartElement("detalles")
        oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecord.DoQuery("exec SP_DetalleFac '" & DocEntry & "','13'")


        If oRecord.RecordCount > 0 Then

            While oRecord.EoF = False
                Dim oRecord2 As SAPbobsCOM.Recordset
                oRecord2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                writer.WriteStartElement("detalle")
                createNode("codigoPrincipal", oRecord.Fields.Item(0).Value.ToString, writer)
                createNode("descripcion", oRecord.Fields.Item(1).Value.ToString, writer)
                createNode("cantidad", oRecord.Fields.Item(2).Value.ToString, writer)
                createNode("precioUnitario", oRecord.Fields.Item(3).Value.ToString, writer)
                createNode("descuento", oRecord.Fields.Item(4).Value.ToString, writer)
                writer.WriteStartElement("impuestos")
                oRecord2.DoQuery("exec SP_Impuesto_Detalle '" & DocEntry & "','" & oRecord.Fields.Item(0).Value.ToString & "','13'")
                If oRecord2.RecordCount > 0 Then
                    While oRecord2.EoF = False
                        writer.WriteStartElement("impuesto")
                        createNode("codigo", oRecord2.Fields.Item(0).Value.ToString, writer)
                        createNode("codigoPorcentaje", oRecord2.Fields.Item(1).Value.ToString, writer)
                        createNode("tarifa", oRecord2.Fields.Item(3).Value.ToString, writer)
                        createNode("baseImponible", oRecord2.Fields.Item(2).Value.ToString, writer)
                        createNode("valor", oRecord2.Fields.Item(4).Value.ToString, writer)
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
                            SetNewTax(fila(1), fila(3).ToString, SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, Double.Parse(fila(9).ToString), fila(10), fila(11), fila(13).ToString, IIf(fila(7).ToString = "", 0, Double.Parse(fila(7).ToString)))
                        ElseIf fila(8).ToString = "IVA" Then
                            SetNewTax(fila(1), fila(3).ToString, SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_VAT, Double.Parse(fila(9).ToString), fila(10), "1-1-010-10-001", fila(13).ToString, IIf(fila(7).ToString = "", 0, Double.Parse(fila(7).ToString)))
                        End If

                    End If
                Next
                If My.Computer.FileSystem.FileExists(Application.StartupPath & "\ATS\Sustentos.txt") = True And My.Computer.FileSystem.FileExists(Application.StartupPath & "\ATS\Comprobantes.txt") = True Then
                    updateValidValues()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub SetFomsUDO()
        Dim ChildTables As New List(Of String)
        Dim findCols As New List(Of String)
        Try
            ChildTables.Clear()
            ChildTables.Add("DGREMISION")
            findCols.Clear()
            findCols.Add("DocEntry")
            findCols.Add("DocNum")
            UDT_UF.AddUDOForm(Me.oCompany, "GREMISION", "GUIA DE REMISION", "GREMISION", SAPbobsCOM.BoUDOObjType.boud_Document, ChildTables, findCols)
            findCols.Clear()
            findCols.Add("Code")
            findCols.Add("Name")
            findCols.Add("U_PLACA")
            ChildTables.Clear()
            ChildTables.Add("T_GTRANSPLACA")
            UDT_UF.AddUDOForm(Me.oCompany, "T_GTRANSPORTISTA", "TRANSPORTISTA", "T_GTRANSPORTISTA", SAPbobsCOM.BoUDOObjType.boud_MasterData, ChildTables, findCols)

            ChildTables.Clear()
            ChildTables.Add("CPDPOSFE")
            findCols.Clear()
            findCols.Add("Docentry")
            findCols.Add("DocNum")
            UDT_UF.AddUDOForm(Me.oCompany, "CPEPOSFE", "CHEQUE POSFECHADO", "CPEPOSFE", SAPbobsCOM.BoUDOObjType.boud_Document, ChildTables, findCols)

            ChildTables.Clear()            
            findCols.Clear()
            findCols.Add("DocEntry")
            findCols.Add("DocNum")
            UDT_UF.AddUDOForm(Me.oCompany, "GREMISION_M", "GUIA DE REMISION MASIVA", "GREMISION_M", SAPbobsCOM.BoUDOObjType.boud_Document, ChildTables, findCols)

        Catch ex As Exception

        End Try
        
    End Sub

    Private Function tipoFactura(p1 As String) As String
        Dim tipofac As String = ""
        Try
            Dim oRecord As SAPbobsCOM.Recordset
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("")
        Catch ex As Exception

        End Try
        Return tipofac
    End Function

    Private Sub PROBAR()


        Dim oSales As SAPbobsCOM.Documents
        Dim erro As Integer = -1
        Dim ddd As String = ""
        oSales = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
        oSales.CardCode = "C20000"
        oSales.CardName = "Norm Thompson"
        oSales.NumAtCard = "0321555"
        oSales.DocDate = "2017/05/23"
        oSales.DocDueDate = "2017/05/23"


        oSales.Lines.ItemCode = "A00001"
        oSales.Lines.Quantity = "3"
        oSales.Lines.UnitPrice = 23
        oSales.Lines.TaxCode = "IVA"
        oSales.Lines.Add()

        oSales.Lines.ItemCode = "A00005"
        oSales.Lines.Quantity = "5"
        oSales.Lines.UnitPrice = 23
        oSales.Lines.TaxCode = "IVA"
        oSales.Lines.Add()

        oSales.Comments = "prueba daniel moreno"


        erro = oSales.Add()
        If erro <> 0 Then
            MessageBox.Show(oCompany.GetLastErrorDescription)
        End If


        Try
            Dim InPay As SAPbobsCOM.Payments
            'Dim oDownPay As SAPbobsCOM.Documents
            'oDownPay = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            Dim sNewObjCode As String = ""

            oCompany.GetNewObjectCode(sNewObjCode) ' This Gets the key of the last created Transaction for this session

            InPay = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

            ' oDownPay.GetByKey(Convert.ToInt32(sNewObjCode))

            InPay.CardCode = "CN0990004196001"

            InPay.Invoices.DocEntry = 26
            InPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice

            InPay.CreditCards.CreditCard = 1  ' Mastercard = 1 , VISA = 2
            InPay.CreditCards.CardValidUntil = CDate("01/12/2020")
            InPay.CreditCards.CreditCardNumber = "1220" ' Just need 4 last digits
            InPay.CreditCards.CreditSum = 8 ' Total Amount of the Invoice
            InPay.CreditCards.VoucherNum = "1234567" ' Need to give the Credit Card confirmation number.
            InPay.CreditCards.PaymentMethodCode = 1
            InPay.CreditCards.Add()

            InPay.CreditCards.CreditCard = 2  ' Mastercard = 1 , VISA = 2
            InPay.CreditCards.CardValidUntil = CDate("01/12/2020")
            InPay.CreditCards.CreditCardNumber = "1220" ' Just need 4 last digits
            InPay.CreditCards.CreditSum = 8 ' Total Amount of the Invoice
            InPay.CreditCards.VoucherNum = "1234567" ' Need to give the Credit Card confirmation number.
            InPay.CreditCards.PaymentMethodCode = 2
            If InPay.Add() <> 0 Then
                MsgBox(oCompany.GetLastErrorDescription())
            Else
                MsgBox("Incoming payment Created!")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

End Class
