Imports SAPbouiCOM
Imports System.Xml
Imports EXO_UIAPI.EXO_UIAPI
Public Class INICIO
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaDatos()
            GenerarParametros()
        End If
    End Sub
    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""
        Dim sSQL As String = ""
        If objGlobal.refDi.comunes.esAdministrador Then
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_WTQ1.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_WTQ1", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If
    End Sub
    Private Sub GenerarParametros()
        If objGlobal.refDi.comunes.esAdministrador Then
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("EXO_NCampo") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("EXO_NCampo", "Attr2Val")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("EXO_Valores") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("EXO_Valores", "'Picking', 'Almacén - Bulk'")
            End If
        End If
    End Sub

    Public Overrides Function filtros() As Global.SAPbouiCOM.EventFilters
        Dim fXML As String = ""
        Try
            fXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml")
            Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
            filtro.LoadFromXML(fXML)
            Return filtro
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        Finally

        End Try
    End Function

    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Dim res As Boolean = True
        Dim Clase As Object = Nothing

        Try
            Select Case infoEvento.FormTypeEx
                Case "180"
                    Clase = New EXO_ORDN(objGlobal)
                    Return CType(Clase, EXO_ORDN).SBOApp_ItemEvent(infoEvento)
                Case "143"
                    Clase = New EXO_OPDN(objGlobal)
                    Return CType(Clase, EXO_OPDN).SBOApp_ItemEvent(infoEvento)
                Case "940"
                    Clase = New EXO_OWTR(objGlobal)
                    Return CType(Clase, EXO_OWTR).SBOApp_ItemEvent(infoEvento)
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
