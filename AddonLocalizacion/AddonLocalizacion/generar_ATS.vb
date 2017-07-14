Public Class generar_ATS
    Private XmlForm As String = Replace(Application.StartupPath & "\ATS.srf", "\\", "\")
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

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "fRtn") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("fRtn")
                oForm.Left = 400
                oForm.Visible = True
                oForm.PaneLevel = 1

                Dim inicio As SAPbouiCOM.EditText
                Dim fin As SAPbouiCOM.EditText
                Dim cmdenviar As SAPbouiCOM.Button
                'esto es para poder hacer que los textos tengan formato de fecha
                oForm.DataSources.DataTables.Add("MyDataTable")
                oForm.DataSources.UserDataSources.Add("Date", SAPbouiCOM.BoDataType.dt_DATE)
                oForm.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_DATE)
                'inicio = oForm.Items.Item("Item_0").Specific
                'fin = oForm.Items.Item("Item_1").Specific
                ' inicio.DataBind.SetBound(True, "", "Date")
                'fin.DataBind.SetBound(True, "", "Date2")

            Else
                oForm = Me.SBO_Application.Forms.Item("fRtn")                
            End If
            Dim oCombo As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_0").Specific
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly
            Dim oCombo2 As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_4").Specific
            oCombo2.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly
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
        If pVal.FormTypeEx = "60006" And pVal.Before_Action = True Then
            If pVal.ItemUID = "btn1" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim oDe As SAPbouiCOM.ComboBox
                Dim oHasta As SAPbouiCOM.ComboBox
                oDe = oForm.Items.Item("Item_0").Specific
                oHasta = oForm.Items.Item("Item_4").Specific
                If oDe.Value.Trim = "" Or oHasta.Value.Trim = "" Then
                    SBOApplication.SetStatusBarMessage("Debe de seleccionar un rango de fecha", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    BubbleEvent = False
                    Return
                Else
                    'generaRetencionXML(oDe.Value, oHasta.Value.ToString, SBOApplication)
                    Dim ats As New generarATS
                    ats.generarXML(oDe.Value.Trim, oHasta.Value.Trim, "", oCompany, SBOApplication)
                    BubbleEvent = False
                    Return
                End If
            End If
        End If
    End Sub
End Class
