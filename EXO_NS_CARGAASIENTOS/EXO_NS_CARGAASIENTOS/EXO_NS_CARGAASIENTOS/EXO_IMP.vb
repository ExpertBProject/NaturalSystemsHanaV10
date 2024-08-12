Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_IMP
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaCampos()
        End If
        cargamenu()
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UTs_EXO_TMPASIENTO.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UTs_EXO_TMPASIENTO", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
        Try
            If objGlobal.SBOApp.Menus.Exists("EXO-MnImpA") = True Then
                Path = objGlobal.refDi.OGEN.pathGeneral & "\02.Menus"  'objGlobal.compañia.conexionSAP.path & "\02.Menus"
                If Path <> "" Then
                    If IO.File.Exists(Path & "\MnCDOC.png") = True Then
                        objGlobal.SBOApp.Menus.Item("EXO-MnImpA").Image = Path & "\MnCDOC.png"
                    End If
                End If
            End If
        Catch ex As Exception
            objGlobal.SBOApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "EXO-MnCASI"
                        'Cargamos pantalla de gestión.
                        If CargarFormImp() = False Then
                            Exit Function
                        End If
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Function CargarFormImp() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing

        CargarFormImp = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_IMP.srf")

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

            CargarFormImp = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_IMP"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_IMP"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                    If EventHandler_et_DOUBLE_CLICK_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select

                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_IMP"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_IMP"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select

                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_et_DOUBLE_CLICK_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

        EventHandler_et_DOUBLE_CLICK_Before = False

        Try
            If pVal.ActionSuccess = False And pVal.ColUID = "Sel" Then
                oForm.Freeze(True)
                For iRow = 0 To oForm.DataSources.DataTables.Item("DT_DOC").Rows.Count - 1
                    If oForm.DataSources.DataTables.Item("DT_DOC").GetValue("Sel", iRow).ToString = "Y" Then
                        oForm.DataSources.DataTables.Item("DT_DOC").SetValue("Sel", iRow, "N")
                    Else
                        oForm.DataSources.DataTables.Item("DT_DOC").SetValue("Sel", iRow, "Y")
                    End If
                Next
            End If

            EventHandler_et_DOUBLE_CLICK_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sArchivo As String = objGlobal.refDi.OGEN.pathGeneral & "\08.Historico\DOC_CARGADOS\" & objGlobal.compañia.CompanyDB & "\ASIENTOS\"
        Dim sTipoArchivo As String = "" : Dim sNomFICH As String = ""
        Dim sArchivoOrigen As String = ""
        Dim sSQL As String = ""
#End Region
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btn_Carga"
#Region "Generar Asiento"
                    If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere generar los Documentos seleccionados?", 1, "Sí", "No") = 1 Then
                        If ComprobarDoc(oForm, "DT_DOC") = True Then
                            oForm.Items.Item("btn_Carga").Enabled = False
                            'Generamos facturas
                            objGlobal.SBOApp.StatusBar.SetText("Creando documentos ... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            oForm.Freeze(True)
                            If CrearDoc(oForm, "DT_DOC") = False Then
                                Exit Function
                            End If
                            oForm.Freeze(False)
                            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                            oForm.Items.Item("btn_Carga").Enabled = True
                        End If
                    End If
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Fin del proceso de carga de registros seleccionados...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
                Case "btn_Fich"
#Region "Coger y leer fichero"
                    Limpiar_Grid(oForm)
                    sTipoArchivo = "Ficheros CSV|*.csv|Texto|*.txt"
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
                        sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                        sArchivo = sArchivo & sNomFICH
                        'Hacemos copia de seguridad para tratarlo
                        Copia_Seguridad(sArchivoOrigen, sArchivo)
                        'Ahora abrimos el fichero para tratarlo
                        TratarFichero(sArchivo, oForm)
                        oForm.Items.Item("btn_Carga").Enabled = True
                    End If
#End Region
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub TratarFichero(ByVal sArchivo As String, ByRef oForm As SAPbouiCOM.Form)
        Dim myStream As StreamReader = Nothing
        Dim Reader As XmlTextReader = New XmlTextReader(myStream)
        Dim sSQL As String = ""
        Dim sExiste As String = "" ' Para comprobar si existen los datos
        Dim sDelimitador As String = "2"
        Try
            objGlobal.SBOApp.StatusBar.SetText("Cargando datos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oForm.Freeze(True)
            EXO_GLOBALES.TratarFichero_TXT(sArchivo, sDelimitador, oForm, objGlobal.compañia, objGlobal.SBOApp, objGlobal)
            oForm.Freeze(False)

#Region "cargar Grid con los datos leidos"
            'Ahora cargamos el Grid con los datos guardados
            objGlobal.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' as ""Sel"",""U_EXO_DOCENTRY"" ""Nº Int."",""U_EXO_DOCNUM"" ""Asiento"", '     ' as ""Estado"",""U_EXO_ASIENTO"" As ""A. Fichero"", "
            sSQL &= " ""U_EXO_FECHA"" as ""Fecha"", ""U_EXO_CUENTA"" as ""Cuenta"", ""U_EXO_DES"" as ""Descripción"", ""U_EXO_DEBE"" as ""DEBE"", ""U_EXO_HABER"" as ""HABER"", "
            sSQL &= " CAST('' as varchar(254)) as ""Descripción Estado"" "
            sSQL &= " From ""@EXO_TMPASIENTO"" "
            sSQL &= " WHERE ""U_EXO_USR""='" & objGlobal.compañia.UserName & "' "
            sSQL &= " ORDER BY ""U_EXO_DOCNUM"", ""U_EXO_LIN"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
#End Region

            objGlobal.SBOApp.StatusBar.SetText("Se terminado de leer el fichero. Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.SBOApp.MessageBox("Se terminado de leer el fichero. Fin del proceso")
        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub
    Private Sub Limpiar_Grid(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            'Limpiamos grid
            'Borrar tablas temporales por usuario activo
            sSQL = "DELETE FROM ""@EXO_TMPASIENTO"" where ""U_EXO_USR""='" & objGlobal.compañia.UserName & "'  "
            oRs.DoQuery(sSQL)
            'Ahora cargamos el Grid con los datos guardados
            objGlobal.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' as ""Sel"",""U_EXO_DOCENTRY"" ""Nº Int."",""U_EXO_DOCNUM"" ""Asiento"", '     ' as ""Estado"",""U_EXO_ASIENTO"" As ""A. Fichero"", "
            sSQL &= " ""U_EXO_FECHA"" as ""Fecha"", ""U_EXO_CUENTA"" as ""Cuenta"", ""U_EXO_DES"" as ""Descripción"", ""U_EXO_DEBE"" as ""DEBE"", ""U_EXO_HABER"" as ""HABER"", "
            sSQL &= " CAST('' as varchar(254)) as ""Descripción Estado"" "
            sSQL &= " From ""@EXO_TMPASIENTO"" "
            sSQL &= " WHERE ""U_EXO_USR""='" & objGlobal.compañia.UserName & "' "
            sSQL &= " ORDER BY ""U_EXO_DOCNUM"", ""U_EXO_LIN"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Private Function CrearDoc(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String) As Boolean
        CrearDoc = False

#Region "Variables"
        Dim sAsiento As String = "" : Dim bPrimeraVez As Boolean = True
        Dim oOJDT As SAPbobsCOM.JournalEntries = Nothing : Dim sTransId As String = "" : Dim sNumber As String = ""
        Dim iLinea As String = ""
        Dim sEstado As String = "OK" : Dim sEstadoDes As String = ""
#End Region
        Try
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sEstado = "OK" : sEstadoDes = ""
                    If sAsiento <> oForm.DataSources.DataTables.Item(sData).GetValue("A. Fichero", i).ToString Then 'Generamos la cabecera
#Region "Graba"
                        If bPrimeraVez = True Then
                            bPrimeraVez = False
                        Else
                            If oOJDT.Add() <> 0 Then
                                sEstado = "ERROR" : sEstadoDes = objGlobal.compañia.GetLastErrorCode & " / " & objGlobal.compañia.GetLastErrorDescription
                                For l = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                                    If sAsiento = oForm.DataSources.DataTables.Item(sData).GetValue("A. Fichero", l).ToString Then
                                        oForm.DataSources.DataTables.Item(sData).SetValue("Estado", l, sEstado)
                                        oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", l, sEstadoDes)
                                    End If
                                Next
                            Else
                                sTransId = objGlobal.compañia.GetNewObjectKey
                                sNumber = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Number"" FROM ""OJDT"" Where ""TransId""=" & sTransId)
                                sEstado = "OK" : sEstadoDes = "Se ha registrado el Asiento Correctamente."
                                For l = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                                    If sAsiento = oForm.DataSources.DataTables.Item(sData).GetValue("A. Fichero", l).ToString Then
                                        oForm.DataSources.DataTables.Item(sData).SetValue("Estado", l, sEstado)
                                        oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", l, sEstadoDes)

                                        oForm.DataSources.DataTables.Item(sData).SetValue("Nº Int.", l, sTransId)
                                        oForm.DataSources.DataTables.Item(sData).SetValue("Asiento", l, sNumber)
                                    End If
                                Next
                            End If

                        End If
#End Region
                        iLinea = 0
                        sAsiento = oForm.DataSources.DataTables.Item(sData).GetValue("A. Fichero", i).ToString
                        oOJDT = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries), SAPbobsCOM.JournalEntries)
                        oOJDT.ReferenceDate = CDate(oForm.DataSources.DataTables.Item(sData).GetValue("Fecha", i).ToString)
                        oOJDT.TaxDate = CDate(oForm.DataSources.DataTables.Item(sData).GetValue("Fecha", i).ToString)
                        oOJDT.DueDate = CDate(oForm.DataSources.DataTables.Item(sData).GetValue("Fecha", i).ToString)
                        oOJDT.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES
                    End If

                    'Generamos lineas

                    If iLinea <> 0 Then
                        oOJDT.Lines.Add()
                    End If
                    oOJDT.Lines.AccountCode = oForm.DataSources.DataTables.Item(sData).GetValue("Cuenta", i).ToString
                    oOJDT.Lines.Debit = oForm.DataSources.DataTables.Item(sData).GetValue("DEBE", i).ToString
                    oOJDT.Lines.Credit = oForm.DataSources.DataTables.Item(sData).GetValue("HABER", i).ToString
                    oOJDT.Lines.AdditionalReference = oForm.DataSources.DataTables.Item(sData).GetValue("Descripción", i).ToString

                    iLinea += 1
                End If
            Next
#Region "Graba"
            If bPrimeraVez = True Then
                bPrimeraVez = False
            Else
                If oOJDT.Add() <> 0 Then
                    sEstado = "ERROR" : sEstadoDes = objGlobal.compañia.GetLastErrorCode & " / " & objGlobal.compañia.GetLastErrorDescription
                    For l = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                        If sAsiento = oForm.DataSources.DataTables.Item(sData).GetValue("A. Fichero", l).ToString Then
                            oForm.DataSources.DataTables.Item(sData).SetValue("Estado", l, sEstado)
                            oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", l, sEstadoDes)
                        End If
                    Next
                Else
                    sTransId = objGlobal.compañia.GetNewObjectKey
                    sNumber = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Number"" FROM ""OJDT"" Where ""TransId""=" & sTransId)
                    sEstado = "OK" : sEstadoDes = "Se ha registrado el Asiento Correctamente."
                    For l = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                        If sAsiento = oForm.DataSources.DataTables.Item(sData).GetValue("A. Fichero", l).ToString Then
                            oForm.DataSources.DataTables.Item(sData).SetValue("Estado", l, sEstado)
                            oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", l, sEstadoDes)

                            oForm.DataSources.DataTables.Item(sData).SetValue("Nº Int.", l, sTransId)
                            oForm.DataSources.DataTables.Item(sData).SetValue("Asiento", l, sNumber)
                        End If
                    Next
                End If

            End If
#End Region

            CrearDoc = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function

    Private Function ComprobarDoc(ByRef oForm As SAPbouiCOM.Form, ByVal sFra As String) As Boolean
        Dim bLineasSel As Boolean = False

        ComprobarDoc = False

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sFra).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sFra).GetValue("Sel", i).ToString = "Y" Then
                    bLineasSel = True
                    Exit For
                End If
            Next

            If bLineasSel = False Then
                objGlobal.SBOApp.MessageBox("Debe seleccionar al menos una línea.")
                Exit Function
            End If

            ComprobarDoc = bLineasSel

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub FormateaGrid(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            For i = 1 To 10
                CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                If i = 8 Or i = 9 Then
                    oColumnTxt.RightJustified = True
                ElseIf i = 1 Then
                    oColumnTxt.LinkedObjectType = 30
                End If
            Next
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Private Sub Copia_Seguridad(ByVal sArchivoOrigen As String, ByVal sArchivo As String)
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
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Comienza la Copia de seguridad del fichero - " & sArchivoOrigen & " -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        If objGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
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
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Copia de Seguridad realizada Correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
End Class

