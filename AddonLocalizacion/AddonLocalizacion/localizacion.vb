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

            If pVal.FormTypeEx = "134" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.BeforeAction = False Then
                Try

                    SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "", "")
                Catch ex As Exception
                    SBO_Application.MessageBox("Operación requiere que muestre campos de usuario", vbOK, "Ok")
                End Try

                'oBusinessForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                'Dim lab As SAPbouiCOM.StaticText
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
                Dim f As SAPbouiCOM.Form
                Dim s As SAPbouiCOM.EditText
                f = SBO_Application.Forms.GetForm("-134", 0)
                s = f.Items.Item("U_DOCUMENTO").Specific
                If s.Value.ToString = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de llenar el campo de Documento", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    BubbleEvent = False
                End If
            Catch ex As Exception

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
    End Sub

    Private Sub SetNewItems()
        Try
            UDT_UF.userField(oCompany, "OCRD", "TIPO IDENTIFICACION", 45, "IDENTIFICACION", SAPbobsCOM.BoFieldTypes.db_Alpha, True, SBO_Application)
            UDT_UF.userField(oCompany, "OCRD", "NO. DOCUMENTO", 45, "DOCUMENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userTable(oCompany, "INF_TRIBUTARIA", "INFORMACION TRIBUTARIA", 45, "NULL", SAPbobsCOM.BoUTBTableType.bott_NoObject, False, SBOApplication)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "AMBIENTE", 11, "AMBIENTE", SAPbobsCOM.BoFieldTypes.db_Numeric, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "EMISION", 11, "EMISION", SAPbobsCOM.BoFieldTypes.db_Numeric, False, SBO_Application)
            'UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "EMISION", 11, "EMISION", SAPbobsCOM.BoFieldTypes.db_Numeric, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "RAZON SOCIAL", 250, "RAZON_SOCIAL", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "NOMBRE COMERCIAL", 250, "NOMBRE_COMERCIAL", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "ESTABLECIMIENTO", 45, "ESTABLECIMIENTO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "PTO EMISOR", 11, "PTO_EMISOR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "DIRECCION", 250, "DIRECCION", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "@INF_TRIBUTARIA", "RUC", 14, "RUC", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            UDT_UF.userField(oCompany, "OINV", "ESTADO", 3, "ESTADO", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
            UDT_UF.userField(oCompany, "OINV", "FIRMA", 60, "FIRMA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBOApplication)
        Catch ex As Exception
            ex.Message.ToString()
        End Try
    End Sub

End Class
