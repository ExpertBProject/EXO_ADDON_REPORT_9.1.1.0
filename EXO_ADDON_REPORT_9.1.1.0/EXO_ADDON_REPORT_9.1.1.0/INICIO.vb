Imports SAPbouiCOM
Imports System.Xml
Imports EXO_UIAPI.EXO_UIAPI
Public Class INICIO
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "Variables"
    Public Shared _sCodeAPANEL As String = "" : Public Shared _sNameAPANEL As String = ""
#End Region
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaDatos()
        End If
        cargamenu()
    End Sub
    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_APANEL.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_APANEL", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults

        If objGlobal.SBOApp.Menus.Exists("EXO-MnHERR") = True Then
            Path = objGlobal.path
            Path = objGlobal.pathMenus   'objGlobal.compañia.conexionSAP.path & "\02.Menus"
            If Path <> "" Then
                If IO.File.Exists(Path & "\MnHEXO.jpg") = True Then
                    objGlobal.SBOApp.Menus.Item("EXO-MnHERR").Image = Path & "\MnHEXO.jpg"
                Else
                    'Sino existe lo copiamos y asignamos
                    EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), "MnHEXO.jpg", Path & "\MnHEXO.jpg")

                    objGlobal.SBOApp.Menus.Item("EXO-MnHERR").Image = Path & "\MnHEXO.jpg"
                End If
            End If
        End If

    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function

    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim Clase As Object = Nothing

        Try

            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.MenuUID
                    Case "1281", "1282" 'Buscar y añadir
                        Clase = New EXO_APANEL(objGlobal)
                        Return CType(Clase, EXO_APANEL).SBOApp_MenuEvent(infoEvento)
                End Select

                'Select Case objGlobal.SBOApp.Forms.ActiveForm.TypeEx
                '    Case "UDO_FT_EXO_CPTOBOAT"
                '        Clase = New EXO_CPTOBOAT(objGlobal)
                '        Return CType(Clase, EXO_CPTOBOAT).SBOApp_MenuEvent(infoEvento)
                'End Select
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnEPA"
                        Clase = New EXO_APANEL(objGlobal)
                        Return CType(Clase, EXO_APANEL).SBOApp_MenuEvent(infoEvento)
                    Case "EXO-MnEAD"
                        Clase = New EXO_ADDON(objGlobal)
                        Return CType(Clase, EXO_ADDON).SBOApp_MenuEvent(infoEvento)
                    Case "EXO-MnERPT"
                        Clase = New EXO_REPORT(objGlobal)
                        Return CType(Clase, EXO_REPORT).SBOApp_MenuEvent(infoEvento)
                End Select
                'Select Case objGlobal.SBOApp.Forms.ActiveForm.TypeEx
                '    Case "UDO_FT_EXO_OBOAT"
                '        Clase = New EXO_OBOAT(objGlobal)
                '        Return CType(Clase, EXO_OBOAT).SBOApp_MenuEvent(infoEvento)
                'End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            Clase = Nothing

        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Dim res As Boolean = True
        Dim Clase As Object = Nothing

        Try
            Select Case infoEvento.FormTypeEx
                Case "UDO_FT_EXO_APANEL"
                    Clase = New EXO_APANEL(objGlobal)
                    Return CType(Clase, EXO_APANEL).SBOApp_ItemEvent(infoEvento)
                Case "EXO_ADDON"
                    Clase = New EXO_ADDON(objGlobal)
                    Return CType(Clase, EXO_ADDON).SBOApp_ItemEvent(infoEvento)
                Case "EXO_REPORT"
                    Clase = New EXO_REPORT(objGlobal)
                    Return CType(Clase, EXO_REPORT).SBOApp_ItemEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Finally
            Clase = Nothing
        End Try
    End Function
End Class
