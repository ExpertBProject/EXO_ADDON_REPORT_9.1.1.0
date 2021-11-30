Imports SAPbobsCOM
Imports SAPbouiCOM
Public Class EXO_APANEL
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub


    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        Dim sExiste As String = ""
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.MenuUID
                    Case "1281", "1282" 'Buscar y añadir
                        oForm = objGlobal.SBOApp.Forms.ActiveForm()
                        If oForm.TypeEx = "UDO_FT_EXO_APANEL" Then
                            Return False
                        End If
                End Select
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnEPA"
                        INICIO._sCodeAPANEL = "INTERCOMPANY"
                        INICIO._sNameAPANEL = "Transacciones entre empresas"
                        'Si no existe, creamos el IC
                        sExiste = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Code"" FROM ""@EXO_APANEL"" WHERE ""Code""='" & INICIO._sCodeAPANEL & "' ")

                        If sExiste = "" Then
                            INICIO._sCodeAPANEL = "INTERCOMPANY"
                            INICIO._sNameAPANEL = "Transacciones entre empresas"
                            'Presentamos UDO Y escribimos los datos de la cabecera
                            objGlobal.funcionesUI.cargaFormUdoBD("EXO_APANEL")
                        Else
                            INICIO._sCodeAPANEL = ""
                            INICIO._sNameAPANEL = ""
                            'Presentamos la pantalla los los datos                              
                            objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_APANEL", "INTERCOMPANY")
                        End If

                End Select
            End If

            Return True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            'Apaño por un error que da EXO_Basic.dll al consultar infoEvento.FormTypeEx
            Try
                If infoEvento.FormTypeEx <> "" Then

                End If
            Catch ex As Exception
                Return False
            End Try

            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_APANEL"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_APANEL"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_APANEL"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(objGlobal, infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_APANEL"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                End If
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Visible = False

                CargarCombos(objGlobal, oForm)
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then 'Para que el combo enseñe la descripción
                    If objGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
                        objGlobal.SBOApp.ActivateMenuItem("1304")
                    End If
                End If

                oItem = oForm.Items.Item("0_U_E")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                oItem = oForm.Items.Item("1_U_E")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)



                If INICIO._sCodeAPANEL <> "" Then
                    oForm.Mode = BoFormMode.fm_ADD_MODE
                    oForm.DataSources.DBDataSources.Item("@EXO_APANEL").SetValue("Code", 0, INICIO._sCodeAPANEL)
                    oForm.DataSources.DBDataSources.Item("@EXO_APANEL").SetValue("Name", 0, INICIO._sNameAPANEL)
                End If



            End If

            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oItem, Object))
        End Try
    End Function
    Private Sub CargarCombos(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim sUsuario_Owner As String = ""
        Dim oRecordSet As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try

            'sUsuario_Owner = objGlobal.refDi.OGEN.usuarioSQL
            'sSQL = "SELECT ""SCHEMA_NAME"" ""BBDD"",""SCHEMA_NAME"" ""SCHEMA"" FROM SYS.SCHEMAS WHERE ""SCHEMA_OWNER""='" & sUsuario_Owner & "' ORDER BY ""SCHEMA_NAME"" "
            'objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").ValidValues, sSQL)

            oRecordSet = objGlobal.compañia.GetCompanyList()
            Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
            Dim oNodes As System.Xml.XmlNodeList = Nothing
            Dim oNode As System.Xml.XmlNode = Nothing
            oXml.LoadXml(oRecordSet.GetAsXML())
            oNodes = oXml.SelectNodes("//row")

            'Añadimos valores
            If oRecordSet.RecordCount > 0 Then
                For j As Integer = 0 To oNodes.Count - 1
                    oNode = oNodes.Item(j)
                    Try
                        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").ValidValues.Add(oNode.SelectSingleNode("dbName").InnerText, oNode.SelectSingleNode("cmpName").InnerText)
                    Catch ex As Exception

                    End Try
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "1" Then
                If pVal.ActionSuccess = True Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then 'Después de añadir
                        objGlobal.SBOApp.ActivateMenuItem("1289")
                    End If
                End If
            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
End Class
