Imports System.Globalization

Public Class Inf_pago_retencion
    Private XmlForm As String = Replace(Application.StartupPath & "\Inf_pago_retencion.srf", "\\", "\")
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

            If UDT_UF.ActivateFormIsOpen(SBO_Application, "I_Pago") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("I_Pago")
                oForm.Left = 400
                oForm.Visible = True
                oForm.PaneLevel = 1
                oForm.DataSources.DataTables.Add("MyDataTable")
            Else
                oForm = Me.SBO_Application.Forms.Item("I_Pago")
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
            If pVal.FormUID = "I_Pago" Then

                If pVal.Before_Action = True And pVal.FormUID = "I_Pago" And pVal.ItemUID = "1" Then
                    Dim ofecha As SAPbouiCOM.EditText = oForm.Items.Item("Item_1").Specific
                    Dim oNumTarjeta As SAPbouiCOM.EditText = oForm.Items.Item("Item_3").Specific
                    Dim oVaucher As SAPbouiCOM.EditText = oForm.Items.Item("Item_5").Specific
                    Dim oComentarios As SAPbouiCOM.EditText = oForm.Items.Item("Item_8").Specific

                    If ofecha.Value.Trim = "" Or oNumTarjeta.Value = "" Or oVaucher.Value = "" Then
                        SBOApplication.SetStatusBarMessage("Debe de Ingresar fecha,numero de tarjeta y voucher ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                        Return
                    End If
                    UDT_UF.infoPago = New Info_pago
                    UDT_UF.infoPago.creditNum = oNumTarjeta.Value.Trim                    
                    UDT_UF.infoPago.validDate = DateTime.ParseExact(ofecha.Value, "yyyyMMdd", CultureInfo.InvariantCulture)
                    UDT_UF.infoPago.remarks = oComentarios.Value
                    UDT_UF.infoPago.voucher = oVaucher.Value
                    oForm.Close()
                    BubbleEvent = False
                    Return
                End If

            End If
        Catch ex As Exception
            SBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try

    End Sub
End Class
