Public Class EXO_CONEXIONES
#Region "Connect to Company"
    Public Shared Sub Connect_Company(ByRef oCompanyDes As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef sUser As String, ByRef sPass As String, ByVal sBBDD As String)

        Try
            'Conectar DI SAP
            oCompanyDes = New SAPbobsCOM.Company
            oCompanyDes.language = SAPbobsCOM.BoSuppLangs.ln_Spanish
            oCompanyDes.Server = oObjGlobal.refDi.compañia.Server 'oCompOrigen.Server
            oCompanyDes.LicenseServer = oObjGlobal.refDi.compañia.LicenseServer ' oCompOrigen.LicenseServer
            oCompanyDes.UserName = sUser
            oCompanyDes.Password = sPass
            oCompanyDes.UseTrusted = False
            oCompanyDes.DbPassword = oObjGlobal.refDi.SQL.claveSQL
            oCompanyDes.DbUserName = oObjGlobal.refDi.SQL.usuarioSQL
            oCompanyDes.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            ' oCompany.SLDServer = oCompOrigen.SLDServer
            oCompanyDes.CompanyDB = sBBDD
            'oLog.escribeMensaje("database:" & oCompany.CompanyDB, EXO_Log.EXO_Log.Tipo.advertencia)
            If oCompanyDes.Connect <> 0 Then
                Throw New System.Exception("Error en la conexión a la compañia:" & oCompanyDes.GetLastErrorDescription.Trim)
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException

            Throw New System.Exception("Error en la conexión a la compañia:" & oCompanyDes.GetLastErrorDescription.Trim & " Error: " & exCOM.Message.ToString)
        Catch ex As Exception
            Throw New System.Exception("Error en la conexión a la compañia:" & oCompanyDes.GetLastErrorDescription.Trim & " Error: " & ex.Message.ToString)

        Finally

        End Try
    End Sub
    Public Shared Sub Disconnect_Company(ByRef oCompany As SAPbobsCOM.Company)
        Try
            If Not oCompany Is Nothing Then
                If oCompany.Connected = True Then
                    oCompany.Disconnect()
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oCompany IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCompany)
            oCompany = Nothing
        End Try
    End Sub

#End Region
End Class
