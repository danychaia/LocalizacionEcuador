Public Class transportista
    Private XmlForm As String = Replace(Application.StartupPath & "\transportista.srf", "\\", "\")
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter
    Public oMatrix As SAPbouiCOM.Matrix
    Public oLineSelected As Integer = -1
    Public Sub New()
        MyBase.New()
        Try
            Me.SBO_Application = UDT_UF.SBOApplication
            Me.oCompany = UDT_UF.Company

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "T_GTRANSPORTISTA_") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("T_GTRANSPORTISTA_")
                oForm.Left = 400
                oForm.Visible = True
                oForm.PaneLevel = 1
                oMatrix = oForm.Items.Item("Item_6").Specific

            Else
                oForm = Me.SBO_Application.Forms.Item("T_GTRANSPORTISTA_")
            End If

            oForm.EnableMenu("1292", True)
            oForm.EnableMenu("1293", True)
            oForm.EnableMenu("1283", False)
            oForm.EnableMenu("1284", False)
            oMatrix.ClearSelections()
            oMatrix.FlushToDataSource()
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
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                oMatrix.Clear()
                oMatrix.AddRow(1)
                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            End If
            If pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                oMatrix.AddRow(1)
                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                BubbleEvent = False
                Return
            End If
            If pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                oForm.Freeze(True)
                If oMatrix.RowCount = 0 Then
                    oMatrix.AddRow(1)
                End If
                oLineSelected = -1
                oForm.Freeze(False)
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
