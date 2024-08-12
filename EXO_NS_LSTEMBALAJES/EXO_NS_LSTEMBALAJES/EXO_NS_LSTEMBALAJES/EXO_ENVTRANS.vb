Imports System.IO
Imports System.Xml
Imports SAPbobsCOM
Imports SAPbouiCOM
Public Class EXO_ENVTRANS
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        SBOApp_MenuEvent = False
        Dim sSQL As String = ""
        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.MenuUID
                    Case "1286" 'Cerrar
                        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.ActiveForm
                        If oForm IsNot Nothing Then
                            If oForm.TypeEx = "UDO_FT_EXO_ENVTRANS" Then
                                If Cerrar_ENVIO(oForm) = False Then
                                    Return False
                                Else
                                    Return True
                                End If
                            End If
                        End If
                End Select
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnGETR"
                        If CargarUDO() = False Then
                            Exit Function
                        End If
                    Case "1282"
                        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.ActiveForm
                        If oForm IsNot Nothing Then
                            If oForm.TypeEx = "UDO_FT_EXO_ENVTRANS" Then
                                EXO_GLOBALES.Modo_Anadir(oForm, objGlobal)
                                Cargar_Combos(oForm)

                                If objGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
                                    objGlobal.SBOApp.ActivateMenuItem("1304")
                                End If
                            End If
                        End If

                End Select
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally

        End Try
    End Function
    Public Function Cerrar_ENVIO(ByRef oform As SAPbouiCOM.Form) As Boolean
#Region "Variables"
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sAlmacen As String = "" : Dim sStatus As String = ""
        Dim sDocEntryFinal As String = "" : Dim sDocNumFinal As String = ""
        Dim dtDatos As System.Data.DataTable = Nothing

        Dim Omercancias As SAPbobsCOM.Documents = Nothing
        Dim sDocEntryCerrar As String = "" : Dim sDocNumCerrar As String = "" : Dim sStatusCerrar As String = ""
        Dim iLinea As Integer = 0

        Dim oGeneralService As SAPbobsCOM.GeneralService = Nothing
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams = Nothing
        Dim oCompService As SAPbobsCOM.CompanyService = objGlobal.compañia.GetCompanyService()
#End Region
        Cerrar_ENVIO = False
        Try
            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            objGlobal.compañia.StartTransaction()

            sDocEntry = oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("DocEntry", 0)
            If sDocEntry = "" Then
                sDocEntry = CType(oform.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value.ToString()
            End If
            sDocNum = oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("DocNum", 0)
            If sDocNum = "" Then
                sDocNum = CType(oform.Items.Item("1_U_E").Specific, SAPbouiCOM.EditText).Value.ToString()
            End If
            sAlmacen = oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("U_EXO_ALMACEN", 0)
            If sAlmacen = "" Then
                sAlmacen = CType(oform.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            End If

            objGlobal.SBOApp.StatusBar.SetText("Cerrando Documento Nº " & sDocNum & "...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sStatus = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Status"" FROM ""@EXO_ENVTRANS"" Where ""DocEntry""=" & sDocEntry)
            If sStatus.Trim = "O" Then
                sSQL = "SELECT TC.""DocEntry"", TC.""DocNum"", COUNT(""U_EXO_IDBULTO"" || ' - ' ||  ""U_EXO_TBULTO"") ""Cantidad"", MAX(""U_EXO_IDBULTO"") ""ID BULTO"", MAX(""U_EXO_TBULTO"") ""BULTO""
                    FROM ""@EXO_LSTEMBL"" TL
                    INNER JOIN ""@EXO_LSTEMB"" TC ON TC.""DocEntry""=TL.""DocEntry""
                    WHERE TC.""Status""='O' and TL.""DocEntry"" IN (SELECT T0.""DocEntry""
           				                        FROM ""@EXO_LSTEMB""  T0 
            			                        Left Join  OCRD T1 ON T0.""U_EXO_IC"" = T1.""CardCode"" 
            			                        where T0.""U_EXO_IDENVIO"" =" & sDocEntry & ")
                    GROUP BY TC.""DocEntry"", TC.""DocNum"", ""U_EXO_IDBULTO"" || ' - ' ||  ""U_EXO_TBULTO""
                    ORDER BY TC.""DocEntry"", MAX(""U_EXO_TBULTO""),MAX(""U_EXO_IDBULTO"")"
                dtDatos = New System.Data.DataTable
                dtDatos = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                If dtDatos.Rows.Count > 0 Then
                    Omercancias = objGlobal.compañia.GetBusinessObject(BoObjectTypes.oInventoryGenExit)
                    Omercancias.DocDate = Date.Now
                    Omercancias.TaxDate = Date.Now
                    For Each MiDataRow As DataRow In dtDatos.Rows
                        If sDocEntryCerrar <> MiDataRow("DocEntry").ToString Then
                            sDocEntryCerrar = MiDataRow("DocEntry").ToString
                            sSQL = "SELECT ""DocNum"" FROM ""@EXO_LSTEMB"" WHERE ""DocEntry""=" & sDocEntryCerrar
                            sDocNumCerrar = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                            sSQL = "SELECT ""Status"" FROM ""@EXO_LSTEMB"" WHERE ""DocEntry""=" & sDocEntryCerrar
                            sStatusCerrar = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                            If sStatusCerrar = "O" Then
                                objGlobal.SBOApp.StatusBar.SetText("Cerrando Lista de embalaje Nº: " & sDocNumCerrar, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
                                'Cerramos el UDO
                                'Get a handle to the SM_MOR UDO
                                oGeneralService = oCompService.GetGeneralService("EXO_LSTEMB")
                                'Close UDO record
                                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                                oGeneralParams.SetProperty("DocEntry", sDocEntryCerrar)
                                oGeneralService.Close(oGeneralParams)
                                objGlobal.SBOApp.StatusBar.SetText("Se ha cerrado la Lista de Embalaje Nº: " & sDocNumCerrar, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)
                            End If
#Region "Lineas"
                            sSQL = "SELECT L.* FROM ""@EXO_PAQL"" L INNER JOIN ""@EXO_PAQ"" C ON C.""Code""=L.""Code"" WHERE C.""Name""='" & MiDataRow("BULTO").ToString & "' ORDER BY L.""LineId"" "
                            oRs.DoQuery(sSQL)
                            For i = 0 To oRs.RecordCount - 1
                                If iLinea <> 0 Then
                                    Omercancias.Lines.Add()
                                End If
                                Omercancias.Lines.ItemCode = oRs.Fields.Item("U_EXO_ITEMCODE").Value.ToString
                                Omercancias.Lines.Quantity = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, MiDataRow("Cantidad").ToString) * EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oRs.Fields.Item("U_EXO_CANT").Value.ToString)
                                Omercancias.Lines.WarehouseCode = sAlmacen
                                ' Omercancias.Lines.BatchNumbers.BatchNumber = ""
                                'Omercancias.Lines.BatchNumbers.Quantity = Omercancias.Lines.Quantity
                                ' Omercancias.Lines.BatchNumbers.Add()
                                Omercancias.Lines.UserFields.Fields.Item("U_EXO_ENVTRDE").Value = sDocEntryCerrar
                                Omercancias.Lines.UserFields.Fields.Item("U_EXO_ENVTRDN").Value = sDocNumCerrar
                                iLinea += 1

                                oRs.MoveNext()
                            Next
#End Region
                        End If

                    Next
                    Omercancias.Comments = "Generado automáticamente al cerrar Envío - Transporte Nº " & sDocNum
                    If Omercancias.Add() <> 0 Then
                        objGlobal.SBOApp.StatusBar.SetText("Error al generar Salida de Mercancía. " & objGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        If objGlobal.compañia.InTransaction = True Then
                            objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        Return False
                    Else
                        objGlobal.compañia.GetNewObjectCode(sDocEntryFinal)

                        If objGlobal.compañia.InTransaction = True Then
                            objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If
                        sSQL = "Select ""DocNum"" FROM """ & objGlobal.compañia.CompanyDB & """.""OIGE"" WHERE ""DocEntry"" = " & sDocEntryFinal
                        oRs.DoQuery(sSQL)
                        If oRs.RecordCount > 0 Then
                            sDocNumFinal = oRs.Fields.Item("DocNum").Value.ToString
                            'Actualizamos el UDO
                            sSQL = "UPDATE ""@EXO_ENVTRANS"" SET ""U_EXO_CONEMB""='" & sDocEntryFinal & "' WHERE ""DocEntry""=" & sDocEntry
                            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                objGlobal.SBOApp.StatusBar.SetText("Se ha generado la Salida de mercancía Nº: " & sDocNumFinal, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                Return True
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("No se ha podido actualizar el envío - Transporte Nº: " & sDocNum & " con el Nº de documento generado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Return False
                            End If

                        Else
                            sDocNumFinal = "0"
                            If objGlobal.compañia.InTransaction = True Then
                                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                            objGlobal.SBOApp.StatusBar.SetText("No se encuentra la Salida de mercancía con Nº Interno: " & sDocEntryFinal, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If

                    End If
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado Lista de embalajes para cerrar. Revise los datos. ", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
                    objGlobal.SBOApp.MessageBox("No se ha encontrado Lista de embalajes para cerrar. Revise los datos. ")
                    If objGlobal.compañia.InTransaction = True Then
                        objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    Return False
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("Este documento ya está cerrado.", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox("Este documento ya está cerrado.")
                If objGlobal.compañia.InTransaction = True Then
                    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If
                Return False
            End If


        Catch ex As Exception
            Throw ex
        Finally
            If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then 'Para que el combo enseñe la descripción
                If objGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
                    objGlobal.SBOApp.ActivateMenuItem("1304")
                End If
            End If

            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(Omercancias, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompService, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oGeneralParams, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oGeneralService, Object))
        End Try
    End Function
    Public Function CargarUDO() As Boolean
        CargarUDO = False

        Try
            objGlobal.funcionesUI.cargaFormUdoBD("EXO_ENVTRANS")

            CargarUDO = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally

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
                        Case "UDO_FT_EXO_ENVTRANS"
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
                        Case "UDO_FT_EXO_ENVTRANS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                    'If EventHandler_MATRIX_LINK_PRESSED(infoEvento) = False Then
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_ENVTRANS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(objGlobal, infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    'If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                                    If EventHandler_GOT_FOCUS_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_ENVTRANS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    'If EventHandler_Choose_FromList_Before(infoEvento) = False Then
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

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
    Private Function EventHandler_GOT_FOCUS_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_GOT_FOCUS_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "22_U_Cb"
                    If oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_FIND_MODE Then
                        If pVal.ItemChanged = True Then
                            Cargar_Combos(oForm)
                        End If
                    End If
            End Select

            EventHandler_GOT_FOCUS_After = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_Before(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "1_U_FD"
                    If oForm.Mode = BoFormMode.fm_OK_MODE Then
                        Rellena_Grid(oForm)
                    ElseIf oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_UPDATE_MODE Then
                        objGlobal.SBOApp.StatusBar.SetText("Grabe primero para poder ver las Expediciones.", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox("Grabe primero para poder ver las expediciones.")
                    End If
            End Select

            EventHandler_ItemPressed_Before = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Sub Rellena_Grid(ByRef oform As SAPbouiCOM.Form)
#Region "Variables"
        Dim sSQL As String = ""
        Dim sDocEntry As String = ""

#End Region
        Try
            sDocEntry = oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("DocEntry", 0)
            If sDocEntry = "" Then
                sDocEntry = CType(oform.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value.ToString()
            End If
            objGlobal.SBOApp.StatusBar.SetText("Documento Nº " & sDocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = "SELECT T0.""U_EXO_IC"" ""Interlocutor"", T1.""CardName"" ""Nombre"", T0.""DocEntry"" , T0.""U_EXO_IDENVIO"" "
            sSQL &= " FROM ""@EXO_LSTEMB""  T0 "
            sSQL &= " Left Join  OCRD T1 ON T0.""U_EXO_IC"" = T1.""CardCode"" "
            sSQL &= " where T0.""U_EXO_IDENVIO"" =" & sDocEntry
            oform.DataSources.DataTables.Item("DTEX").ExecuteQuery(sSQL)
            FormateaGrid(oform)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub FormateaGrid(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)

            For i = 0 To 3
                Select Case i
                    Case 2
                        CType(oform.Items.Item("grdEX").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdEX").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                        oColumnTxt.LinkedObjectType = "EXO_LSTEMB"
                    Case Else
                        CType(oform.Items.Item("grdEX").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdEX").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next
            CType(oform.Items.Item("grdEX").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Function EventHandler_ItemPressed_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "1"
                    If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                        Cargar_Combos(oForm)
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
        Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim sFecha As String = ""
        Dim sAlmacen As String = ""
#End Region

        EventHandler_COMBO_SELECT_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True And oForm.Mode = BoFormMode.fm_ADD_MODE Then
                If pVal.ItemUID = "4_U_Cb" Then
                    If CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value IsNot Nothing Then
                        Dim sSerie As String = CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                        Dim iNum As Integer
                        iNum = oForm.BusinessObject.GetNextSerialNumber(CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, oForm.BusinessObject.Type.ToString)
                        oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("DocNum", 0, iNum.ToString)
                    End If
                End If
            End If
            If oForm.Visible = True Then
                If pVal.ItemUID = "22_U_Cb" Then ' Almacen
                    sFecha = CType(oForm.Items.Item("21_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                    If sFecha = "" Then
                        sFecha = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                        oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_DOCDATE", 0, sFecha)
                    End If
                    sAlmacen = CType(oForm.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    'Expedición
                    sSQL = "Select ""TrnspCode"",""TrnspName"" FROM OSHP WHERE ""Active""='Y' and ""TrnspCode"" in ("
                    sSQL &= " SELECT distinct  ""TrnspCode"" FROM ("
                    sSQL &= " Select T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM ORDR T0 "
                    sSQL &= " Inner JOIN RDR1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    sSQL &= " UNION ALL "
                    sSQL &= " Select T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM ODLN T0 "
                    sSQL &= " Inner JOIN DLN1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    sSQL &= " UNION ALL "
                    sSQL &= "Select  T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM OPRR T0 "
                    sSQL &= " Inner JOIN PRR1 t1 on  T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    sSQL &= " UNION ALL "
                    sSQL &= " Select T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM  ORPD T0 "
                    sSQL &= " Inner JOIN RPD1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    sSQL &= " UNION ALL "
                    sSQL &= "Select  T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM OWTQ T0 "
                    sSQL &= " Inner JOIN WTQ1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' )"
                    sSQL &= " ) ORDER BY ""TrnspName"""
                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                ElseIf pVal.ItemUID = "20_U_Cb" Then 'Clase de expedición
                    If CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value IsNot Nothing Then
                        Dim sExpedicion As String = CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                        sSQL = "Select IFNULL(""U_EXO_AGE"",'') FROM OSHP WHERE ""TrnspCode""='" & sExpedicion & "' "
                    End If
                    Dim sAgeCod As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_AGTCODE", 0, sAgeCod)
                    sSQL = "SELECT ""CardName"" FROM OCRD WHERE ""CardCode""='" & sAgeCod & "' "
                    Dim sAgeNom As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_AGTNAME", 0, sAgeNom)
                End If
            End If

            EventHandler_COMBO_SELECT_After = True

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
    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True And oForm.TypeEx = "UDO_FT_EXO_ENVTRANS" Then
                'No dejamos que modifique la cabecera
                oItem = oForm.Items.Item("0_U_E")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                oItem = oForm.Items.Item("22_U_Cb")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem = oForm.Items.Item("4_U_Cb")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem = oForm.Items.Item("20_U_Cb")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem = oForm.Items.Item("21_U_E")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                Cargar_Combos(oForm)

                If objGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
                    objGlobal.SBOApp.ActivateMenuItem("1304")
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
    Private Sub Cargar_Combos(ByRef oform As SAPbouiCOM.Form)
#Region "Variables"
        Dim sClaseExp As String = ""
        Dim sSucursal As String = ""
        Dim sSerieDef As String = ""
        Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim sFecha As String = ""
        Dim sSQL As String = ""
        Dim sAlmacendef As String = ""
        Dim sExpedicion As String = ""
#End Region
        Try
            sFecha = CType(oform.Items.Item("21_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
            If sFecha = "" Then
                sFecha = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_DOCDATE", 0, sFecha)
            End If

            'Almacen
            sSQL = "SELECT ""Branch"" FROM OUSR WHERE ""USERID""=" & objGlobal.compañia.UserSignature.ToString
            sSucursal = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            sSQL = " SELECT ""WhsCode"",""WhsName"" FROM OWHS"
            sSQL &= " WHERE ""Inactive""='N' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            If oform.Mode = BoFormMode.fm_ADD_MODE Then
                'Poner almacen por defecto
                Try
                    sSQL = " SELECT TOP 1 ""WhsCode"" FROM OWHS"
                    sSQL &= " WHERE ""Inactive""='N'  "
                    sAlmacendef = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    CType(oform.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(sAlmacendef, BoSearchKey.psk_ByValue)
                Catch ex As Exception

                End Try
            Else
                sAlmacendef = oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("U_EXO_ALMACEN", 0)
            End If
            oform.Items.Item("22_U_Cb").DisplayDesc = True

            'Expedición
            If objGlobal.compañia.DbServerType = BoDataServerTypes.dst_HANADB Then
                sSQL = " SELECT '-1' ""TrnspCode"", '' ""TrnspName"" FROM DUMMY "
            Else
                sSQL = " SELECT '-1' ""TrnspCode"", '' ""TrnspName"" "
            End If
            sSQL &= " UNION ALL "
            sSQL &= "SELECT ""TrnspCode"",""TrnspName"" FROM OSHP WHERE ""Active""='Y' "
            sSQL &= " ORDER BY ""TrnspName"""
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)


            'Series 
            sSQL = "SELECT ""Series"",""SeriesName"" FROM NNM1 WHERE ""ObjectCode""='EXO_ENVTRANS' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oform.Items.Item("4_U_Cb").DisplayDesc = True



            If oform.Mode = BoFormMode.fm_ADD_MODE Then
                'Expedición por defecto
                If CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues.Count > 0 Then
                    CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)
                    sExpedicion = CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                Else
                    sExpedicion = ""
                End If

                'Poner serie por defecto y el num. de documento
                sSQL = " SELECT ""DfltSeries"" FROM ONNM WHERE ""ObjectCode""='EXO_ENVTRANS' "
                sSerieDef = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                CType(oform.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(sSerieDef, BoSearchKey.psk_ByValue)

                EXO_GLOBALES.Poner_DocNum(oform, sSerieDef, objGlobal)

                ''Como en la expedición tenemos la agencia, pues tenemos que rellenarlo automático
                'sSQL = "SELECT IFNULL(""U_EXO_AGE"",'') FROM OSHP WHERE ""TrnspCode""='" & sExpedicion & "' "
                'Dim sAgeCod As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                'oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_AGTCODE", 0, sAgeCod)
                'sSQL = "SELECT ""CardName"" FROM OCRD WHERE ""CardCode""='" & sAgeCod & "' "
                'Dim sAgeNom As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                'oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_AGTNAME", 0, sAgeNom)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    '    Private Sub Cargar_Combo_Matricula_Conductor_Plataforma(ByRef oform As SAPbouiCOM.Form, ByVal sAgencia As String)
    '#Region "Variables"
    '        Dim sSQL As String = ""

    '#End Region
    '        Try

    '            'Matricula
    '            sSQL = "SELECT ""U_EXO_VEHICULO"",""U_EXO_DES"" FROM ""@EXO_VEHIAGL"" WHERE ""Code""='" & sAgencia & "' "
    '            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("25_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
    '            oform.Items.Item("25_U_Cb").DisplayDesc = True

    '            'Conductor 
    '            sSQL = "SELECT ""U_EXO_COND"",(""U_EXO_NOMBRE"" || ' ' || ""U_EXO_APE"") ""Nombre"" FROM ""@EXO_CONAGL"" WHERE ""Code""='" & sAgencia & "' "
    '            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("26_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
    '            oform.Items.Item("26_U_Cb").DisplayDesc = True

    '            'Plataforma
    '            sSQL = "SELECT ""U_EXO_PLATA"",""U_EXO_PLATAD"" FROM ""@EXO_PLATAAGL"" WHERE ""Code""='" & sAgencia & "' "
    '            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_5").ValidValues, sSQL)
    '            CType(oform.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_5").DisplayDesc = True

    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Sub
    'Private Function EventHandler_Choose_FromList_Before(ByVal pVal As ItemEvent) As Boolean
    '    Dim oCFLEvento As IChooseFromListEvent = Nothing
    '    Dim oForm As SAPbouiCOM.Form = Nothing
    '    Dim oConds As SAPbouiCOM.Conditions = Nothing
    '    Dim oCond As SAPbouiCOM.Condition = Nothing
    '    Dim oRs As SAPbobsCOM.Recordset = Nothing
    '    Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
    '    Dim oNodes As System.Xml.XmlNodeList = Nothing
    '    Dim oNode As System.Xml.XmlNode = Nothing

    '    EventHandler_Choose_FromList_Before = False

    '    Try
    '        If pVal.ItemUID = "23_U_E" Then 'Agencia de transporte
    '            oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
    '            oCFLEvento = CType(pVal, IChooseFromListEvent)

    '            oConds = New SAPbouiCOM.Conditions
    '            oCond = oConds.Add
    '            oCond.Alias = "QryGroup1"
    '            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCond.CondVal = "Y"
    '            'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

    '            oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
    '        End If

    '        EventHandler_Choose_FromList_Before = True

    '    Catch ex As Exception
    '        objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
    '    Finally
    '        EXO_CleanCOM.CLiberaCOM.Form(oForm)
    '    End Try
    'End Function
    'Private Function EventHandler_Choose_FromList_After(ByVal pVal As ItemEvent) As Boolean
    '    Dim oCFLEvento As IChooseFromListEvent = Nothing
    '    Dim oDataTable As DataTable = Nothing
    '    Dim oForm As SAPbouiCOM.Form = Nothing
    '    Dim oRs As SAPbobsCOM.Recordset = Nothing
    '    Dim sSQL As String = ""
    '    Dim sNombre As String = ""
    '    EventHandler_Choose_FromList_After = False

    '    Try
    '        oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
    '        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
    '            oForm = Nothing
    '            Return True
    '        End If

    '        oCFLEvento = CType(pVal, IChooseFromListEvent)

    '        oDataTable = oCFLEvento.SelectedObjects
    '        If Not oDataTable Is Nothing Then
    '            Select Case oCFLEvento.ChooseFromListUID
    '                Case "CFLAT"
    '                    oDataTable = oCFLEvento.SelectedObjects

    '                    If oDataTable IsNot Nothing Then
    '                        If pVal.ItemUID = "23_U_E" Then
    '                            CType(oForm.Items.Item("24_U_E").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardName", 0).ToString
    '                            Cargar_Combo_Matricula_Conductor_Plataforma(oForm, oDataTable.GetValue("CardCode", 0).ToString)
    '                        End If
    '                    End If
    '                'Case "CFLIC"
    '                '    oDataTable = oCFLEvento.SelectedObjects

    '                '    If oDataTable IsNot Nothing Then
    '                '        If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_1" Then
    '                '            sNombre = oDataTable.GetValue("CardName", 0).ToString
    '                '            oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANSEX").SetValue("U_EXO_EMPRESA", pVal.Row - 1, sNombre)
    '                '        End If
    '                '    End If
    '                Case "CFLAB"
    '                    oDataTable = oCFLEvento.SelectedObjects

    '                    If oDataTable IsNot Nothing Then
    '                        If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_1_1" Then
    '                            sNombre = oDataTable.GetValue("Name", 0).ToString
    '                            oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANSAB").SetValue("U_EXO_AGNAME", pVal.Row - 1, sNombre)
    '                        End If
    '                    End If
    '            End Select
    '        End If

    '        EventHandler_Choose_FromList_After = True

    '    Catch ex As Exception
    '        objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
    '    Finally
    '        EXO_CleanCOM.CLiberaCOM.FormDatatable(oDataTable)
    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
    '        EXO_CleanCOM.CLiberaCOM.Form(oForm)
    '    End Try
    'End Function
    Public Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument
        Dim sFecha As String = "" : Dim sAlmacen As String = "" : Dim sSucursal As String = ""
        Dim sSQL As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_ENVTRANS"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Mode = BoFormMode.fm_OK_MODE Then
                                    'Almacen
                                    'sSQL = "SELECT ""Branch"" FROM OUSR WHERE ""USERID""=" & objGlobal.compañia.UserSignature.ToString
                                    'sSucursal = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                    sSQL = " SELECT ""WhsCode"",""WhsName"" FROM OWHS"
                                    sSQL &= " WHERE ""Inactive""='N' "
                                    'ssql &= " And ""U_EXO_SUCURSAL""='" & sSucursal & "' "
                                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                                    ''Expedición
                                    'sSQL = "Select ""TrnspCode"",""TrnspName"" FROM OSHP "
                                    'objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)


                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If Comprobar_existe(oForm) = False Then
                                    Return False
                                Else
                                    Return True
                                End If
                        End Select

                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_ENVTRANS"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
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
    Private Function Comprobar_existe(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Comprobar_existe = False
        Dim sClaseExp As String = "" : Dim sAlmacen As String = ""
        Dim sFecha As String = "" : Dim sDocNum As String = "" : Dim sSerie As String = ""
        Dim sSQL As String = ""
        Try
            sSerie = CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            sFecha = CType(oForm.Items.Item("21_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
            sClaseExp = CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            sAlmacen = CType(oForm.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            sSQL = " SELECT ""DocNum"" FROM ""@EXO_ENVTRANS"" Where ""U_EXO_CEXP""='" & sClaseExp & "' and ""U_EXO_DOCDATE""='" & sFecha & "' and ""U_EXO_ALMACEN""='" & sAlmacen & "' and ""Series""=" & sSerie
            sDocNum = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            If sDocNum = "" Then
                Return True
            Else
                objGlobal.SBOApp.StatusBar.SetText("Ya existe el documento Nº " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If
            Comprobar_existe = True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
