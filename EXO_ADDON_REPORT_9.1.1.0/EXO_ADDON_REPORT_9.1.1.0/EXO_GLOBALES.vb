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
    Public Shared Function Import_Report(ByRef oCompanyDes As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI,
                                         ByVal sNOMFICH As String, ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Varibales"
        Dim oLayoutService As SAPbobsCOM.ReportLayoutsService = Nothing
        Dim oReport As SAPbobsCOM.ReportLayout = Nothing
        Dim sTypeCode As String = ""
        Dim oCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oBlobParams As SAPbobsCOM.BlobParams = Nothing
        Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment = Nothing
        Dim oBlob As SAPbobsCOM.Blob = Nothing
        Dim sReportExiste As String = ""
        Dim sSQL As String = ""
        Dim bPonerLayoutDFLT As Boolean = False : Dim sReportDFLT As String = ""
#End Region
        Import_Report = False
        Try
            oLayoutService = CType(oCompanyDes.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), SAPbobsCOM.ReportLayoutsService)
            oReport = CType(oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout), SAPbobsCOM.ReportLayout)

            'Initialize critical properties 
            ' Use TypeCode "RCRI" to specify a Crystal Report. 
            ' Use other TypeCode to specify a layout for a document type. 
            ' List of TypeCode types are in table RTYP. 
            sTypeCode = oForm.DataSources.UserDataSources.Item("UDF").Value.ToString
            Select Case sTypeCode
                Case "I" : sTypeCode = "RCRI"
                Case Else : sTypeCode = oForm.DataSources.UserDataSources.Item("UDL").Value.ToString
            End Select
            oReport.Name = oForm.DataSources.UserDataSources.Item("UDNOM").Value.ToString
            oReport.TypeCode = sTypeCode
            oReport.Author = oCompanyDes.UserName
            oReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal
            oReport.Localization = "ES"

            Dim newReportCode As String = ""
            Try
                ' Add New object 
                oReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal

                'Comprobamos si Existe
                sSQL = "SELECT ""DocCode"" FROM  """ & oCompanyDes.CompanyDB & """.""RDOC"" WHERE ""DocName""='" & oForm.DataSources.UserDataSources.Item("UDNOM").Value.ToString & "' "
                sReportExiste = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                If sReportExiste <> "" Then
                    'Comprobar si es por defecto si TYPE CODE<> RCRI
                    If sTypeCode <> "RCRI" Then
                        sSQL = "SELECT ""DEFLT_REP"" FROM """ & oCompanyDes.CompanyDB & """.""RTYP"" WHERE ""CODE""='" & oForm.DataSources.UserDataSources.Item("UDL").Value.ToString & "' "
                        sReportDFLT = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                        If sReportDFLT = sReportExiste Then
                            bPonerLayoutDFLT = True
                            sSQL = "UPDATE """ & oCompanyDes.CompanyDB & """.""RTYP"" SET ""DEFLT_REP""='' WHERE ""CODE""='" & oForm.DataSources.UserDataSources.Item("UDL").Value.ToString & "' "
                            oObjGlobal.refDi.SQL.sqlUpdB1(sSQL)
                        End If
                    End If

                    Dim oExisteReportParams As SAPbobsCOM.ReportLayoutParams = CType(oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams), SAPbobsCOM.ReportLayoutParams)
                    oExisteReportParams.LayoutCode = sReportExiste
                    oLayoutService.DeleteReportLayout(oExisteReportParams)
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Se borra Report / Layaout existente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If

                Dim oNewReportParams As SAPbobsCOM.ReportLayoutParams = CType(oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams), SAPbobsCOM.ReportLayoutParams)
                Select Case sTypeCode
                    Case "RCRI" : oNewReportParams = oLayoutService.AddReportLayoutToMenu(oReport, oForm.DataSources.UserDataSources.Item("UDL").Value.ToString)
                    Case Else : oNewReportParams = oLayoutService.AddReportLayout(oReport)
                End Select

                'Get code of the added ReportLayout object 
                newReportCode = oNewReportParams.LayoutCode
                If sReportDFLT <> "" And bPonerLayoutDFLT = True Then
                    sSQL = "UPDATE """ & oCompanyDes.CompanyDB & """.""RTYP"" SET ""DEFLT_REP""='" & newReportCode & "' WHERE ""CODE""='" & oForm.DataSources.UserDataSources.Item("UDL").Value.ToString & "' "
                    oObjGlobal.refDi.SQL.sqlUpdB1(sSQL)
                End If
            Catch ex As Exception
                Dim sError As String = Err.Description
                oObjGlobal.SBOApp.StatusBar.SetText(sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

            ' Wpload .rpt file using SetBlob interface 
            Dim rptFilePath As String = sNOMFICH

            oCompanyService = oCompanyDes.GetCompanyService()
            'Specify the table And field to update 
            oBlobParams = CType(oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams), SAPbobsCOM.BlobParams)
            oBlobParams.Table = "RDOC"
            oBlobParams.Field = "Template"

            ' Specify the record whose blob field Is to be set 
            oKeySegment = oBlobParams.BlobTableKeySegments.Add()
            oKeySegment.Name = "DocCode"
            oKeySegment.Value = newReportCode

            oBlob = CType(oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob), SAPbobsCOM.Blob)

            ' Put the rpt file into buffer 
            Dim oFile As FileStream = New FileStream(rptFilePath, System.IO.FileMode.Open)
            Dim fileSize As Integer = CType(oFile.Length, Integer)
            Dim buf(CInt(oFile.Length() - 1)) As Byte
            oFile.Read(buf, 0, fileSize)
            oFile.Close()


            ' Convert memory buffer to Base64 string 
            oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)

            Try
                'Upload Blob to database 
                oCompanyService.SetBlob(oBlobParams, oBlob)
            Catch ex As Exception
                Dim sError As String = Err.Description
                oObjGlobal.SBOApp.StatusBar.SetText(sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

            Import_Report = True
        Catch ex As Exception
            Throw ex
        Finally
#Region "Liberar"
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oReport, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oLayoutService, Object))

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlob, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oKeySegment, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlobParams, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompanyService, Object))

#End Region
        End Try
    End Function
    Public Shared Function Import_Report2(ByRef oCompanyDes As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI,
                                         ByVal sNOMFICH As String, ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Varibales"
        Dim oLayoutService As SAPbobsCOM.ReportLayoutsService = Nothing
        Dim oReport As SAPbobsCOM.ReportLayout = Nothing
        Dim sTypeCode As String = ""
        Dim oCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oBlobParams As SAPbobsCOM.BlobParams = Nothing
        Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment = Nothing
        Dim oBlob As SAPbobsCOM.Blob = Nothing
        Dim sReportExiste As String = ""
        Dim sSQL As String = ""
#End Region
        Import_Report2 = False
        Try
            oLayoutService = CType(oCompanyDes.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), SAPbobsCOM.ReportLayoutsService)
            oReport = CType(oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout), SAPbobsCOM.ReportLayout)

            'Initialize critical properties 
            ' Use TypeCode "RCRI" to specify a Crystal Report. 
            ' Use other TypeCode to specify a layout for a document type. 
            ' List of TypeCode types are in table RTYP. 
            sTypeCode = oForm.DataSources.UserDataSources.Item("UDF").Value.ToString
            Select Case sTypeCode
                Case "I" : sTypeCode = "RCRI"
                Case Else : sTypeCode = oForm.DataSources.UserDataSources.Item("UDL").Value.ToString
            End Select
            oReport.Name = oForm.DataSources.UserDataSources.Item("UDNOM").Value.ToString
            oReport.TypeCode = sTypeCode
            oReport.Author = oCompanyDes.UserName
            oReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal
            oReport.Localization = "ES"

            Dim newReportCode As String = ""
            Try
                ' Add New object 
                oReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal

                'Comprobamos si Existe y borramos
                sSQL = "SELECT ""DocCode"" FROM  """ & oCompanyDes.CompanyDB & """.""RDOC"" WHERE ""DocName""='" & oForm.DataSources.UserDataSources.Item("UDNOM").Value.ToString & "' "
                sReportExiste = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                If sReportExiste <> "" Then
                    Dim oExisteReportParams As SAPbobsCOM.ReportLayoutParams = CType(oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams), SAPbobsCOM.ReportLayoutParams)
                    oExisteReportParams.LayoutCode = sReportExiste
                    oLayoutService.DeleteReportLayout(oExisteReportParams)
                End If

                Dim oNewReportParams As SAPbobsCOM.ReportLayoutParams = CType(oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams), SAPbobsCOM.ReportLayoutParams)
                Select Case sTypeCode
                    Case "RCRI" : oNewReportParams = oLayoutService.AddReportLayoutToMenu(oReport, "12800")
                    Case Else : oNewReportParams = oLayoutService.AddReportLayout(oReport)
                End Select

                'Get code of the added ReportLayout object 
                newReportCode = oNewReportParams.LayoutCode

            Catch ex As Exception
                Dim sError As String = Err.Description
                oObjGlobal.SBOApp.StatusBar.SetText(sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

            ' Wpload .rpt file using SetBlob interface 
            Dim rptFilePath As String = sNOMFICH

            oCompanyService = oCompanyDes.GetCompanyService()
            'Specify the table And field to update 
            oBlobParams = CType(oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams), SAPbobsCOM.BlobParams)
            oBlobParams.Table = "RDOC"
            oBlobParams.Field = "Template"

            ' Specify the record whose blob field Is to be set 
            oKeySegment = oBlobParams.BlobTableKeySegments.Add()
            oKeySegment.Name = "DocCode"
            oKeySegment.Value = newReportCode

            oBlob = CType(oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob), SAPbobsCOM.Blob)

            ' Put the rpt file into buffer 
            Dim oFile As FileStream = New FileStream(rptFilePath, System.IO.FileMode.Open)
            Dim fileSize As Integer = CType(oFile.Length, Integer)
            Dim buf(CInt(oFile.Length() - 1)) As Byte
            oFile.Read(buf, 0, fileSize)
            oFile.Close()


            ' Convert memory buffer to Base64 string 
            oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)

            Try
                'Upload Blob to database 
                oCompanyService.SetBlob(oBlobParams, oBlob)
            Catch ex As Exception
                Dim sError As String = Err.Description
                oObjGlobal.SBOApp.StatusBar.SetText(sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

            Import_Report2 = True
        Catch ex As Exception
            Throw ex
        Finally
#Region "Liberar"
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oReport, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oLayoutService, Object))

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlob, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oKeySegment, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlobParams, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompanyService, Object))

#End Region
        End Try
    End Function
End Class
