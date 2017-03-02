Public Class guia_remision

    Private XmlForm As String = Replace(Application.StartupPath & "\guia_remision.xml", "\\", "\")
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

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "GREMISION_") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("GREMISION_")
                oForm.Left = 400
                oForm.Visible = True
                oForm.PaneLevel = 1

            Else
                oForm = Me.SBO_Application.Forms.Item("GREMISION_")
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
        If FormUID = "GREMISION_" Then
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
                    If (pVal.ItemUID = "Item_9") Then
                        Try
                            Dim txtRuc As SAPbouiCOM.EditText = oForm.Items.Item("Item_9").Specific
                            ' Dim txtNomEmp As SAPbouiCOM.EditText = oForm.Items.Item("1000008").Specific
                            val = oDataTable.GetValue("CardCode", 0)
                            txtRuc.Value = val
                        Catch ex As Exception

                        End Try
                                            
                    End If
                    If (pVal.ItemUID = "Item_28") Then
                        Try
                            Dim oDoc As SAPbouiCOM.EditText = oForm.Items.Item("Item_28").Specific
                            ' Dim txtNomEmp As SAPbouiCOM.EditText = oForm.Items.Item("1000008").Specific
                            val = oDataTable.GetValue("DocEntry", 0)
                            oDoc.Value = val                            
                        Catch ex As Exception

                        End Try

                    End If
                End If
                
            End If
        End If
    End Sub
    Private Sub carcarSeries()

    End Sub


End Class
