Imports System.IO
Imports OfficeOpenXml
Imports SAPbouiCOM
Public Class EXO_REPORT
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        Dim sExiste As String = ""
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            If infoEvento.BeforeAction = True Then
                Dim ORS As SAPbobsCOM.Recordset = objGlobal.compañia.GetCompanyList

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnERPT"
                        If CargarFormREPORT() = False Then
                            Return False
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
    Public Function CargarFormREPORT() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oDI_CONFAC As EXO_DIAPI.EXO_UDOEntity = Nothing
        Dim sCodeEnvio As String = ""

        CargarFormREPORT = False

        Try

            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_REPORT.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try
            CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Select("I", BoSearchKey.psk_ByValue)
            Carga_Datos_Menu(oForm)
            CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Item.Enabled = False
            oForm.Visible = True

            CargarFormREPORT = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Visible = True

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oFP, Object))
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
                        Case "EXO_REPORT"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_After(infoEvento) = False Then
                                        Return False
                                    End If
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
                        Case "EXO_REPORT"
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
                        Case "EXO_REPORT"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_REPORT"
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
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim sBBDD As String = "" : Dim sUser As String = "" : Dim sPass As String = ""
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sTipoArchivo As String = ""
        Dim sArchivoOrigen As String = ""
        Dim sArchivo As String = ""
        Dim sNomFICH As String = ""
        Dim OdtEmpresas As System.Data.DataTable = Nothing : Dim oCompanyDes As SAPbobsCOM.Company = Nothing
        Dim sSQL As String = ""
#End Region

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "btn_Carga" Then
#Region "Cargar REPORT"
                sArchivo = objGlobal.path & "\05.Rpt\"
                sArchivoOrigen = CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value.ToString
                If sArchivoOrigen.Trim <> "" Then
                    sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                    sArchivo = sArchivo & sNomFICH
                    'Hacemos copia de seguridad para tratarlo
                    OdtEmpresas = New System.Data.DataTable
                    OdtEmpresas.Clear()
                    sSQL = "SELECT * FROM ""@EXO_IPANELL"" WHERE ""Code""='INTERCOMPANY' and ""U_EXO_TIPO""='D' "
                    OdtEmpresas = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                    If OdtEmpresas.Rows.Count > 0 Then
                        objGlobal.SBOApp.StatusBar.SetText("Se va a proceder a recorrer las SOCIEDADES...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                        For Each dr As DataRow In OdtEmpresas.Rows
                            Try
                                sBBDD = dr.Item("U_EXO_BBDD").ToString : sUser = dr.Item("U_EXO_USER").ToString : sPass = dr.Item("U_EXO_PASS").ToString
                                EXO_CONEXIONES.Connect_Company(oCompanyDes, objGlobal, sUser, sPass, sBBDD)
                                objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Importando Report: " & sNomFICH, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                sSQL = "SELECT ""U_EXO_PATH"" FROM """ & oCompanyDes.CompanyDB & """.""@EXO_OGEN"""
                                sArchivo = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                sArchivo &= "\05.Rpt\"
                                sArchivo = sArchivo & sNomFICH
                                EXO_GLOBALES.Copia_Seguridad(objGlobal, sArchivoOrigen, sArchivo)
                                'Importamos el report
                                EXO_GLOBALES.Import_Report(oCompanyDes, objGlobal, sArchivo, oForm)

                            Catch ex As Exception
                                objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Error: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Finally
                                objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Fin Sincronización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                EXO_CONEXIONES.Disconnect_Company(oCompanyDes)
                            End Try
                        Next

                    Else
                        EXO_GLOBALES.Copia_Seguridad(objGlobal, sArchivoOrigen, sArchivo)
                    End If
                    objGlobal.SBOApp.MessageBox(" Fin Sincronización.")
                Else
                    objGlobal.SBOApp.StatusBar.SetText("Sin fichero no se puede importar el Report...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                End If
#End Region
            ElseIf pVal.ItemUID = "btn_Fich" Then
#Region "Coger la ruta del fichero"
                Dim sFormato As String = "" : Dim sLayout As String = ""
                sFormato = CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                If CType(oForm.Items.Item("cbLayout").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                    sLayout = CType(oForm.Items.Item("cbLayout").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                Else
                    sLayout = ""
                End If

                If sFormato <> "" Then
                    Select Case sLayout
                        Case ""
                            If sFormato = "L" Then
                                objGlobal.SBOApp.MessageBox("Debe indicar dónde importar el Layout.")
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Debe indicar dónde importar el Layout.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                oForm.Items.Item("btn_Carga").Enabled = False
                                Exit Function
                            End If
                    End Select
                    If oForm.DataSources.UserDataSources.Item("UDNOM").Value.ToString = "" Then
                        objGlobal.SBOApp.MessageBox("Debe indicar el nombre del Report.")
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) -Debe indicar el nombre del Report.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        oForm.Items.Item("btn_Carga").Enabled = False
                        Exit Function
                    End If
                    sTipoArchivo = "RPT|*.rpt"

                    'Tenemos que controlar que es cliente o web
                    If objGlobal.SBOApp.ClientType = SAPbouiCOM.BoClientType.ct_Browser Then
                        sArchivoOrigen = objGlobal.SBOApp.GetFileFromBrowser() 'Modificar
                    Else
                        'Controlar el tipo de fichero que vamos a abrir según campo de formato
                        sArchivoOrigen = objGlobal.funciones.OpenDialogFiles("Abrir archivo como", sTipoArchivo)
                    End If

                    If Len(sArchivoOrigen) = 0 Then
                        CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = ""
                        objGlobal.SBOApp.MessageBox("Debe indicar un archivo a importar.")
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Debe indicar un archivo a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        oForm.Items.Item("btn_Carga").Enabled = False
                        Exit Function
                    Else
                        CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = sArchivoOrigen
                        oForm.Items.Item("btn_Carga").Enabled = True
                    End If
                Else
                    objGlobal.SBOApp.MessageBox("No ha seleccionado el formato a importar.")
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - No ha seleccionado el formato a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Active = True
                End If
#End Region
            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            If OdtEmpresas IsNot Nothing Then
                OdtEmpresas = Nothing
            End If

        End Try
    End Function
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim sSQL As String = ""
        Dim oForm As SAPbouiCOM.Form = Nothing
#End Region

        EventHandler_COMBO_SELECT_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "cb_Format" Then
                Select Case CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    Case "I"
                        Carga_Datos_Menu(oForm)
                    Case "L"
                        sSQL = "SELECT ""CODE"" ""Código"", ""NAME"" ""Nombre"" from """ & objGlobal.compañia.CompanyDB & """.""RTYP"" ORDER BY  ""CODE"" "
                        objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbLayout").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                End Select
            End If
            EventHandler_COMBO_SELECT_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally


        End Try
    End Function
    Private Sub Carga_Datos_Menu(ByRef oForm As SAPbouiCOM.Form)
        Dim sPath As String = ""
        Dim sCodigo As String = ""
        Dim sDescripcion As String = ""
        Dim sSQL As String = ""
        Dim pck As ExcelPackage = Nothing
        Dim iLin As Integer = 1 : Dim sContenido As String = "-"

        Dim oVatGroup As SAPbobsCOM.VatGroups = Nothing
        Try
            objGlobal.SBOApp.StatusBar.SetText("Rellenando Lista Menú ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sPath = objGlobal.path
            sPath = objGlobal.refDi.OGEN.rutaConsultas
            If sPath = "" Then
                objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado el Path del fichero a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            Else
                sPath += "\MENU.xlsx"
                If IO.File.Exists(sPath) = False Then
                    'Sino existe lo copiamos y asignamos
                    EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), "MENU.xlsx", sPath)
                End If
            End If

            ' miramos si existe el fichero y cargamos
            If File.Exists(sPath) Then
                Dim excel As New FileInfo(sPath)
                pck = New ExcelPackage(excel)
                Dim workbook = pck.Workbook
                Dim worksheet = workbook.Worksheets.First()
                sSQL = "SELECT ' ' ""Código"", ' ' ""Nombre"" from ""DUMMY"" "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbLayout").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

                While sContenido.Trim <> ""
                    iLin += 1
                    sContenido = worksheet.Cells("A" & iLin).Text
                    If sContenido.Trim <> "" Then
                        sCodigo = worksheet.Cells("A" & iLin).Text
                        sDescripcion = worksheet.Cells("B" & iLin).Text
                        CType(oForm.Items.Item("cbLayout").Specific, SAPbouiCOM.ComboBox).ValidValues.Add(sCodigo, sDescripcion)
                    Else
                        Exit While
                    End If
                End While

                objGlobal.SBOApp.StatusBar.SetText("Se rellenado la lista de menú", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                pck.Dispose()
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado el Path del fichero a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
        End Try
    End Sub
End Class
