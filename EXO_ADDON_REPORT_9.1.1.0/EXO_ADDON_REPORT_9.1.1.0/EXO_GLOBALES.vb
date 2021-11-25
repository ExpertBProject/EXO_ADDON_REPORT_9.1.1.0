Imports System.IO
Imports SAPbouiCOM

Public Class EXO_GLOBALES
    Public Shared Sub CopiarRecurso(ByVal pAssembly As Reflection.Assembly, ByVal pNombreRecurso As String, ByVal pRuta As String)
        Dim s As Stream = pAssembly.GetManifestResourceStream(pAssembly.GetName().Name + "." + pNombreRecurso)
        If s.Length = 0 Then
            Throw New Exception("No se puede encontrar el recurso '" + pNombreRecurso + "'")
        Else
            Dim buffer(CInt(s.Length() - 1)) As Byte
            s.Read(buffer, 0, buffer.Length)

            Dim sw As BinaryWriter = New BinaryWriter(File.Open(pRuta, FileMode.Create))
            sw.Write(buffer)
            sw.Close()
        End If
    End Sub
    Public Shared Sub Copia_Seguridad(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sArchivoOrigen As String, ByVal sArchivo As String)
        'Comprobamos el directorio de copia que exista
        Dim sPath As String = ""
        sPath = IO.Path.GetDirectoryName(sArchivo)
        If IO.Directory.Exists(sPath) = False Then
            IO.Directory.CreateDirectory(sPath)
        End If
        If IO.File.Exists(sArchivo) = True Then
            IO.File.Delete(sArchivo)
        End If
        'Subimos el archivo
        oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Comienza la Copia de seguridad del fichero - " & sArchivoOrigen & " -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        If oObjGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
            Dim fs As FileStream = New FileStream(sArchivoOrigen, FileMode.Open, FileAccess.Read)
            Dim b(CInt(fs.Length() - 1)) As Byte
            fs.Read(b, 0, b.Length)
            fs.Close()
            Dim fs2 As New System.IO.FileStream(sArchivo, IO.FileMode.Create, IO.FileAccess.Write)
            fs2.Write(b, 0, b.Length)
            fs2.Close()
        Else
            My.Computer.FileSystem.CopyFile(sArchivoOrigen, sArchivo)
        End If
        oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Copia de Seguridad realizada Correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub

    Public Shared Function Sincroniza_Addon(ByRef oCompanyDes As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sNOMFICH As String, ByVal sAddon As String) As Boolean
#Region "Varibales"
        Dim oGeneralService As SAPbobsCOM.GeneralService = Nothing
        Dim oGeneralData As SAPbobsCOM.GeneralData = Nothing
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams = Nothing
        Dim sCmp As SAPbobsCOM.CompanyService = Nothing
        Dim oChild As SAPbobsCOM.GeneralData = Nothing
        Dim oChildren As SAPbobsCOM.GeneralDataCollection = Nothing
        Dim sSQL As String = "" : Dim sExiste As String = "" : Dim sOrden As String = ""
#End Region
        Sincroniza_Addon = False
        Try
            sCmp = oCompanyDes.GetCompanyService()
            oGeneralService = sCmp.GetGeneralService("EXO_OGEN")
            ' Get UDO record
            oGeneralParams = CType(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams), SAPbobsCOM.GeneralDataParams)
            oGeneralParams.SetProperty("Code", "EXO_KERNEL")
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)
            oChildren = oGeneralData.Child("EXO_OGEN2")
            'Miramos si existe el Addon
            sSQL = "SELECT ""U_EXO_NAME"" FROM """ & oCompanyDes.CompanyDB & """.""@EXO_OGEN2"" WHERE ""U_EXO_RUTA""='" & sNOMFICH & "' "
            sExiste = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
            If sExiste = "" Then
                oChild = oChildren.Add()
                oChild.SetProperty("U_EXO_NAME", sAddon)
                oChild.SetProperty("U_EXO_INFO", sAddon)
                oChild.SetProperty("U_EXO_RUTA", sNOMFICH)
                oChild.SetProperty("U_EXO_ACT", "Y")
                oChild.SetProperty("U_EXO_UPD", "Y")
                sSQL = "SELECT MAX(U_EXO_ORDEN)+1 FROM """ & oCompanyDes.CompanyDB & """.""@EXO_OGEN2"" "
                sOrden = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                oChild.SetProperty("U_EXO_ORDEN", sOrden)
                oGeneralService.Update(oGeneralData)
            End If

            Sincroniza_Addon = True
        Catch ex As Exception
            Throw ex
        Finally
#Region "Liberar"
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oGeneralService, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oGeneralData, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oGeneralParams, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(sCmp, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oChild, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oChildren, Object))
#End Region
        End Try
    End Function
End Class
