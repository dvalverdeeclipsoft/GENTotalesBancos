Imports System.IO
Imports Utilidades.CtrlArchivos
Imports FuncionesBaseDatos
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Module ModMain
    Private ReadOnly vDbControlOperador As New FuncionesBD
    Private ReadOnly vDbControlPlataforma As New FuncionesBD
    Private ReadOnly vDbControlReporteria As New FuncionesBD
    Private ReadOnly vDbControlProgreso As New FuncionesBD

    Private oCnn As OleDbConnection
    Private cmd As OleDbCommand

    Private swLog As StreamWriter

    '----------------------------
    Dim vgProcCodigo As String
    Dim vgProcEmisor As String = ""
    Dim vgProcXlsHoja As String = ""
    Dim vgProcTipo As String = ""
    Dim vgProcReplicar As String = ""
    Dim vgProcEliminar As String = ""
    Dim vgProcAnio As String = ""
    Dim vgProcMes As String = ""
    Dim vgProcDia As String = ""
    Dim vgProcArchivo As String = ""
    '----------------------------
    Dim vgTotBancosds As DataSet = Nothing
    Dim vgTotIntermBancosds As DataSet = Nothing
    Private vgFechaIni As String
    Private vgFechaFin As String

    Private vgProgresoGlobal As Integer = 0
    Private vgProgresoGlobalAnterior As Integer = 0
    Private vgPorcionElimiacion As Integer
    Private vgPorcionLecturaReg As Integer
    Private vgPorcionGrabarData As Integer
    Private vgRegCargados As Integer
    Private vgRegError As Integer
    Private vgOrigenOl As String()

    Private vParametrosUpdProgreso As New ParametrosCmd
    Private vgAcumulador As New CLAcumulador

    Dim vgItmsOperadoras As New Collection
    Dim vgItmsOperadorasPc As New Collection
    Dim vgItmsTipoCods As New Collection

    Private Enum EConsultas
        coCodigo = 0
        coEmisor = 1
        coUsuario = 2
        coFecha = 3
        coHora = 4
        coFiltros = 5
        coHoraIni = 6
        coHoraFin = 7
        coObservacion = 8
        coEstado = 9
    End Enum

    Sub Main(ByVal cmdArgs() As String)
        Dim vOutWriter As StreamWriter
        Dim vPathConexiones As String = AppDomain.CurrentDomain.BaseDirectory & "\FuncionesBaseDatos.inf"
        Try

            If cmdArgs.Length = 1 Then
                vgProcCodigo = cmdArgs.GetValue(0).ToString
                vOutWriter = New StreamWriter(AppDomain.CurrentDomain.BaseDirectory & "\ErrorLog\GENTotalesBancos_" & vgProcCodigo & "_" & Format(Now, "yyyyMMdd") & "_" & Format(Now, "HHmmssfff") & ".err")
                vOutWriter.AutoFlush = True
                Console.SetOut(vOutWriter)
                PConsola("Iniciando GENTotalesBancos...")
            Else
                vOutWriter = New StreamWriter(AppDomain.CurrentDomain.BaseDirectory & "\ErrorLog\GENTotalesBancos_" & Format(Now, "yyyyMMdd") & "_" & Format(Now, "HHmmssfff") & ".err")
                vOutWriter.AutoFlush = True
                Console.SetOut(vOutWriter)
                PConsola("Iniciando GENTotalesBancos...")
                Return
            End If

            vDbControlOperador.PdbSetArchivoInf(vPathConexiones)
            vDbControlPlataforma.PdbSetArchivoInf(vPathConexiones)
            vDbControlProgreso.PdbSetArchivoInf(vPathConexiones)
            vDbControlReporteria.PdbSetArchivoInf(vPathConexiones)

            vgItmsOperadoras.Add("PORTA", "PORTA")
            vgItmsOperadoras.Add("MOVISTAR", "MOVISTAR")
            vgItmsOperadoras.Add("ALEGRO", "ALEGRO")

            vgItmsOperadorasPc.Add("Porta", "Porta")
            vgItmsOperadorasPc.Add("MOVISTAR", "MOVISTAR")
            vgItmsOperadorasPc.Add("ALEGRO", "ALEGRO")


            'vgItmsTipoCods.Add("LGBVI", "LGBVI")
            'vgItmsTipoCods.Add("LGBVE", "LGBVE")
            'vgItmsTipoCods.Add("NEOBO", "NEOBO")
            'vgItmsTipoCods.Add("VPNBG", "VPNBG")
            'vgItmsTipoCods.Add("EFECM", "EFECM")
            'vgItmsTipoCods.Add("OTPCR", "OTPCR")
            'vgItmsTipoCods.Add("OTPNW", "OTPNW")
            'vgItmsTipoCods.Add("NEOVL", "NEOVL")

            vgOrigenOl = My.Settings.vgOrigenBGOl.Split(CChar(","))
            For Each vOrigen As String In vgOrigenOl.ToList
                vgItmsTipoCods.Add(vOrigen, vOrigen)
            Next


            With vParametrosUpdProgreso
                .Campos.Append("tb_porcentaje")
                .Tipos = "N"
                .Valores.Append("0")
                .Condicion.Append("tb_id = " & vgProcCodigo)
                .Tabla = "aa_totales_bancos_carga"
            End With

            PActualizarProceso()
            PGenerarParametros()
            PInicializaTablaTotales()
            PProcesarCargaDeBancos()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
            PConsola("Error: " & ex.Message)
        Finally
            vDbControlProgreso.FdbCierreConexionSql()
            PConsola("Finalizando GENTotalesBancos...")
        End Try

    End Sub
    Private Sub PActualizarProceso()
        Try
            Dim vParametrosUpd As New ParametrosCmd
            If Not vDbControlOperador.FdbConexionSql(My.Settings.vgServerOperador) Then
                PConsola("No se pudo establecer la conexión con la base de datos de Operadores.")
                Return
            End If
            With vParametrosUpd
                .Campos.Append("tb_estado")
                .Tipos = "C"
                .Valores.Append("PRO")
                .Condicion.Append("tb_id = " & vgProcCodigo)
                .Tabla = "aa_totales_bancos_carga"
            End With
            If Not vDbControlOperador.FdbUpdateSql("app_admin", vParametrosUpd) Then
                PConsola("No se pudo grabar el estado del procesamiento de totales bancos.")
            End If
        Catch ex As Exception
            Return
        Finally
            vDbControlOperador.FdbCierreConexionSql()
        End Try
    End Sub
    Private Sub PConsola(ByVal parError As String)
        Console.WriteLine(Format(Now, "yyyy-MM-dd HH:mm:ss.fff") & Space(1) & parError)
    End Sub
    Private Sub PGenerarParametros()
        Dim vParametrosProceso As String() = {}
        Dim vParametrosSel As New ParametrosCmd
        Dim vListadoPendiente As Object = Nothing

        Try
            'Obtenemos los parámetros del proceso...
            If Not vDbControlOperador.FdbConexionSql(My.Settings.vgServerOperador) Then
                PConsola("No se pudo establecer la conexión con la base de datos.")
                Return
            End If

            With vParametrosSel
                .Campos.Append("tb_parametros_utilizados, tb_archivoruta")
                .Condicion.Append("tb_id = " & vgProcCodigo & " and tb_estado = 'PRO'")
                .Tabla = "aa_totales_bancos_carga"
            End With
            'Determinar consultas por generar
            If Not vDbControlOperador.FdbSelectSql("app_admin", vParametrosSel, vListadoPendiente) Then
                PConsola("El código enviado como parámetro no es válido para esta operación.")
                Return
            End If
            Dim vLstPendiente = DirectCast(vListadoPendiente, List(Of List(Of Object)))
            For Each vItem As List(Of Object) In vLstPendiente
                vParametrosProceso = vItem.Item(0).ToString.Split(Convert.ToChar("&"))
                vgProcArchivo = vItem.Item(1).ToString
            Next

            For Each vItem As String In vParametrosProceso
                Dim vArrDatos As String() = vItem.Split(Convert.ToChar(":"))
                Select Case vArrDatos(0)
                    Case "Emisor"
                        vgProcEmisor = vArrDatos(1)
                    Case "Hoja"
                        vgProcXlsHoja = vArrDatos(1)
                    Case "Tipo"
                        vgProcTipo = vArrDatos(1)
                    Case "Replicacion"
                        vgProcReplicar = vArrDatos(1)
                    Case "EliminarDataAnterior"
                        vgProcEliminar = vArrDatos(1)
                    Case "Año"
                        vgProcAnio = vArrDatos(1)
                    Case "Mes"
                        vgProcMes = vArrDatos(1)
                    Case "Dia"
                        vgProcDia = vArrDatos(1)
                End Select
            Next
        Catch ex As Exception
            PConsola("No se pudo obtener los parámetros del proceso.")
        Finally
            vDbControlOperador.FdbCierreConexionSql()
        End Try
    End Sub
    Private Function FValidarPreviaEliminacion() As Boolean
        If vgProcAnio = "" Then
            PConsola("No se ha especificado el año del cual desea eliminar la información.")
            Return False
        End If
        If vgProcMes = "" Then
            PConsola("No se ha especificado el mes del cual desea eliminar la información.")
            Return False
        End If
        If (vgProcDia = "" AndAlso vgProcTipo = "Diario") OrElse (vgProcDia = "NO APLICA" AndAlso vgProcTipo = "Diario") Then
            PConsola("No se ha especificado el día del cual desea eliminar la información.")
            Return False
        End If
        Return True
    End Function
    Private Sub PProcesarCargaDeBancos()
        Dim vResultado As String = ""
        Try
            If vgProcEliminar.ToUpper = "SI" Then

                If Not FValidarPreviaEliminacion() Then Return

                vgPorcionElimiacion = 5
                vgPorcionLecturaReg = 90
                vgPorcionGrabarData = 5

                PBorrarInfoAnterior()
                PObtieneFecha()
            Else
                vgPorcionElimiacion = 0
                vgPorcionLecturaReg = 95
                vgPorcionGrabarData = 5
                PObtieneFecha1()
            End If

            If vgProcEmisor.ToUpper = "BAMAZONAS" OrElse vgProcEmisor.ToUpper = "BGUAYAQUIL" Then
                If vgProcTipo = "Diario" Then
                    PCargarDatosBancosBG()
                Else
                    PCargarDatosBancos()
                End If
            ElseIf vgProcEmisor.ToUpper = "PACIFICARD" Then
                PCargarDatosBancosPacf()
            End If

            vResultado = "Registros cargados: " & CStr(vgRegCargados) & vbNewLine &
                        "Registros erróneos: " & CStr(vgRegError) & vbNewLine &
                        "Total registros: " & CStr(vgRegCargados + vgRegError)

        Catch ex As Exception
            PConsola("No se pudo cargar la información. Error: " & ex.Message & " Trazabilidad: " & ex.StackTrace)
            vResultado = "ERROR" & vbNewLine &
                        ex.Message & vbNewLine &
                        "RESUMEN" & vbNewLine &
                        "Registros cargados: " & CStr(vgRegCargados) & vbNewLine &
                        "Registros erróneos: " & CStr(vgRegError) & vbNewLine &
                        "Total registros: " & CStr(vgRegCargados + vgRegError)
        Finally
            PFinalizarProceso(vResultado)
        End Try
    End Sub
    Private Sub PFinalizarProceso(ByVal parResumen As String)
        Try
            Dim vParametrosUpd As New ParametrosCmd
            If Not vDbControlOperador.FdbConexionSql(My.Settings.vgServerOperador) Then
                PConsola("No se pudo establecer la conexión con la base de datos de Operadores.")
                Return
            End If
            With vParametrosUpd
                .Campos.Append("tb_porcentaje|tb_estado|tb_resumen")
                .Tipos = "N|C|C"
                .Valores.Append("100|FIN|" & parResumen.Replace("'", ""))
                .Condicion.Append("tb_id = " & vgProcCodigo)
                .Tabla = "aa_totales_bancos_carga"
            End With
            If Not vDbControlOperador.FdbUpdateSql("app_admin", vParametrosUpd) Then
                PConsola("No se pudo grabar el estado final del procesamiento de totales bancos.")
            Else
                PConsola("Fin del proceso.")
            End If
        Catch ex As Exception
            Return
        Finally
            vDbControlOperador.FdbCierreConexionSql()
        End Try
    End Sub
    Private Sub FPActualizaProgreso(ByVal parValorEnCurso As Integer, ByVal parValorTotal As Integer, ByVal parFactor As Integer, ByVal parTipoAcum As ENTipoAcumulador)
        Dim vResultado As Integer
        Try
            vResultado = CInt(Math.Floor(CDbl((parValorEnCurso * parFactor) / parValorTotal)))
            Select Case parTipoAcum
                Case ENTipoAcumulador.ProcesoEliminacion
                    vgAcumulador.propAvanceEliminacion = vResultado
                Case ENTipoAcumulador.ProcesoLectura
                    vgAcumulador.propAvanceLectura = vResultado
                Case ENTipoAcumulador.ProcesoGrabar
                    vgAcumulador.propAvanceGrabarData = vResultado
            End Select

            vgProgresoGlobal = vgAcumulador.propAvanceEliminacion + vgAcumulador.propAvanceLectura + vgAcumulador.propAvanceGrabarData


            If vgProgresoGlobal <> vgProgresoGlobalAnterior Then
                If Not vDbControlProgreso.FdbConexionSql(My.Settings.vgServerOperador) Then
                    PConsola("No se pudo establecer la conexión con la base de datos Transaccional.")
                    Return
                End If
                vParametrosUpdProgreso.Valores = New Text.StringBuilder(CStr(vgProgresoGlobal))
                vDbControlProgreso.FdbUpdateSql("app_admin", vParametrosUpdProgreso)
            End If
        Catch ex As Exception
            Return
        Finally
            vgProgresoGlobalAnterior = vgProgresoGlobal
            vDbControlProgreso.FdbCierreConexionSql()
        End Try
    End Sub
    Private Sub PInicializaTablaTotales()
        Dim vClavesPrimTot(4) As DataColumn
        Dim vClavesPrimTotInterm(5) As DataColumn
        Dim vDtTotales As New System.Data.DataTable("Totales")
        Dim vDtIntermedio As New System.Data.DataTable("TotalesInter")
        Try
            'Tabla para los totales 
            Dim vColumnTot As New DataColumn("Emisor", Type.GetType("System.String"))
            vDtTotales.Columns.Add(vColumnTot)
            vClavesPrimTot(0) = vColumnTot
            vColumnTot = New DataColumn("Shortcode", Type.GetType("System.String"))
            vDtTotales.Columns.Add(vColumnTot)
            vClavesPrimTot(1) = vColumnTot
            vColumnTot = New DataColumn("Tipo", Type.GetType("System.String"))
            vDtTotales.Columns.Add(vColumnTot)
            vClavesPrimTot(2) = vColumnTot
            vColumnTot = New DataColumn("Operadora", Type.GetType("System.String"))
            vDtTotales.Columns.Add(vColumnTot)
            vClavesPrimTot(3) = vColumnTot
            vColumnTot = New DataColumn("Fecha", Type.GetType("System.String"))
            vDtTotales.Columns.Add(vColumnTot)
            vClavesPrimTot(4) = vColumnTot
            vColumnTot = New DataColumn("Total", Type.GetType("System.String"))
            vDtTotales.Columns.Add(vColumnTot)
            vDtTotales.PrimaryKey = vClavesPrimTot
            vDtTotales.AcceptChanges()
            vgTotBancosds = New DataSet("Datos")
            vgTotBancosds.Tables.Add(vDtTotales)

            'Tabla para los totales intermedios
            vColumnTot = New DataColumn("Emisor", Type.GetType("System.String"))
            vDtIntermedio.Columns.Add(vColumnTot)
            vClavesPrimTotInterm(0) = vColumnTot
            vColumnTot = New DataColumn("Shortcode", Type.GetType("System.String"))
            vDtIntermedio.Columns.Add(vColumnTot)
            vClavesPrimTotInterm(1) = vColumnTot
            vColumnTot = New DataColumn("Tipo", Type.GetType("System.String"))
            vDtIntermedio.Columns.Add(vColumnTot)
            vClavesPrimTotInterm(2) = vColumnTot
            vColumnTot = New DataColumn("Operadora", Type.GetType("System.String"))
            vDtIntermedio.Columns.Add(vColumnTot)
            vClavesPrimTotInterm(3) = vColumnTot
            vColumnTot = New DataColumn("Fecha", Type.GetType("System.String"))
            vDtIntermedio.Columns.Add(vColumnTot)
            vClavesPrimTotInterm(4) = vColumnTot
            vColumnTot = New DataColumn("Total", Type.GetType("System.String"))
            vDtIntermedio.Columns.Add(vColumnTot)
            vColumnTot = New DataColumn("Estado", Type.GetType("System.String"))
            vDtIntermedio.Columns.Add(vColumnTot)
            vClavesPrimTotInterm(5) = vColumnTot
            vDtIntermedio.PrimaryKey = vClavesPrimTotInterm
            vDtIntermedio.AcceptChanges()
            vgTotIntermBancosds = New DataSet("Datos")
            vgTotIntermBancosds.Tables.Add(vDtIntermedio)
        Catch ex As Exception
            PConsola("No se pudo inicializar las tablas de los totales generales e intermedios.")
        End Try
    End Sub
    Private Sub PBorrarInfoAnterior()
        Try
            Dim vFecha As String
            Dim vFecha2 As String
            Dim vCondicion As String
            Dim vMes, vMes2 As String
            Dim vAnio As String
            Dim vTablaTotales As String = "sms_totales_bancos"

            vMes = vgProcMes
            vMes2 = (CInt(vMes) + 1).ToString
            vAnio = vgProcAnio
            vFecha = vAnio & "-" & Microsoft.VisualBasic.Right("0" & vMes, 2) & "-" & "01"
            vFecha2 = vAnio & "-" & Microsoft.VisualBasic.Right("0" & vMes2, 2) & "-" & "01"
            Select Case CInt(vMes)
                Case 12
                    vFecha = vAnio & "-12-" & "01"
                    vFecha2 = CStr(CInt(vAnio) + 1) & "-01-" & "01"
            End Select
            If vgProcEmisor = "Diario" Then
                vFecha = vgProcAnio & "-" & vgProcMes & "-" & Microsoft.VisualBasic.Right("0" & vgProcDia, 2)
                vCondicion = "tb_emisor = '" & vgProcEmisor & "' and tb_fecha ='" & vFecha & "'"
            Else
                vCondicion = "tb_emisor = '" & vgProcEmisor & "' and tb_fecha >='" & vFecha & "' and tb_fecha <'" & vFecha2 & "'"
            End If

            If Not vDbControlPlataforma.FdbConexionSql(My.Settings.vgServerPlataforma) Then
                PConsola("No se pudo establecer la conexión con la base de datos de la plataforma transaccional.")
                Return
            End If
            If Not vDbControlReporteria.FdbConexionSql(My.Settings.vgServerReporteria) Then
                PConsola("No se pudo establecer la conexión con la base de datos de reportería.")
                Return
            End If

            If vgProcTipo = "Mensual" AndAlso vgProcReplicar = "No" Then
                vTablaTotales = "sms_totales_bancos_tmp"
            End If

            If Not vDbControlPlataforma.FdbDeleteSql("msgswitchweb", vTablaTotales, vCondicion) Then
                PConsola("No se pudo realizar la elimación en plataforma transaccional. Se continúa con el proceso...")
            End If

            FPActualizaProgreso(2, 5, vgPorcionElimiacion, ENTipoAcumulador.ProcesoEliminacion)

            If vTablaTotales <> "sms_totales_bancos_tmp" Then
                If Not vDbControlPlataforma.FdbDeleteSql("msgswitchweb", "sms_diarios_bancos", vCondicion) Then
                    PConsola("No se pudo realizar la elimación en plataforma transaccional (sms_diarios_bancos). Se continúa con el proceso...")
                End If

                If vDbControlReporteria.FdbIfExistsSql("sms_reportes", "sms_totales_bancos", vCondicion) AndAlso
                   Not vDbControlReporteria.FdbDeleteSql("sms_reportes", "sms_totales_bancos", vCondicion) Then
                    PConsola("No se pudo realizar la elimación en Reportería. Se continúa con el proceso...")
                End If
            End If
            FPActualizaProgreso(5, 5, vgPorcionElimiacion, ENTipoAcumulador.ProcesoEliminacion)
        Catch ex As Exception
            PConsola("No se pudo realizar el proceso de eliminación.")
        Finally
            vDbControlPlataforma.FdbCierreConexionSql()
            vDbControlReporteria.FdbCierreConexionSql()
        End Try
    End Sub
    Private Sub PObtieneFecha()
        Dim vFechaAnt As Date
        Dim vAnio, vMes, vDiaIni, vDiaFin As String

        vFechaAnt = New Date(Now.Year, Now.Month, Now.Day, 0, 0, 0, DateTimeKind.Local)
        vAnio = CStr(vFechaAnt.Year)
        vMes = CStr(vFechaAnt.Month - 1)
        vDiaFin = CStr(Date.DaysInMonth(vFechaAnt.Year, vFechaAnt.Month - 1))

        If (vFechaAnt.Month) >= 10 Then
            vMes = CStr(vFechaAnt.Month)
        Else
            vMes = "0" & CStr(vFechaAnt.Month)
        End If

        vDiaIni = "01"

        vgFechaIni = vAnio & "-" & vMes & "-" & vDiaIni
        vgFechaFin = vAnio & "-" & vMes & "-" & vDiaFin
    End Sub
    'Private Sub PInicializaExcel()

    '    Try
    '        Dim exDet As Exception
    '        vgExcel = New Excel.Application
    '        If vgExcel Is Nothing Then
    '            exDet = New Exception("Error en acceso a Excel.")
    '            Throw exDet
    '        End If

    '        vgLibro = vgExcel.Workbooks.Open(vgProcArchivo)

    '        If vgLibro Is Nothing Then
    '            exDet = New Exception("Error en acceso a archivo Excel.")
    '            Throw exDet
    '        End If
    '    Catch ex As Exception
    '        PConsola(ex.Message)
    '        Dim exGen As Exception
    '        exGen = New Exception(ex.Message)
    '        Throw exGen
    '    End Try
    'End Sub
    'Private Sub PLiberaExcel()
    '    If vgLibro IsNot Nothing Then
    '        vgLibro.Close(False)
    '        vgLibro = Nothing
    '    End If
    '    If Not vgExcel IsNot Nothing Then
    '        vgExcel = Nothing
    '    End If
    '    vgLibro = Nothing
    '    vgExcel = Nothing
    'End Sub
    'Private Sub PLecturaDatosPc()
    '    Dim k As Integer
    '    Dim vShortcode As String
    '    Dim vFecha As String
    '    Dim vTotal As String
    '    Dim vCol2 As String
    '    Dim vCol4 As String
    '    Dim vCol6 As String
    '    Dim vCol5 As String
    '    Dim vCantFilasXls As Integer
    '    Dim vRange As Range

    '    vgHoja = CType(vgLibro.Worksheets(vgProcXlsHoja), Excel.Worksheet)
    '    vgHoja.Activate()
    '    vCantFilasXls = vgHoja.Range("a1").CurrentRegion.Rows.Count
    '    k = 2
    '    vRange = CType(vgHoja.Cells(k, 1), Range)
    '    While Not String.IsNullOrEmpty(Trim(CStr(vRange.Value)))
    '        vRange = CType(vgHoja.Cells(k, 1), Range)
    '        vShortcode = CStr(vRange.Value)
    '        vRange = CType(vgHoja.Cells(k, 7), Range)
    '        vTotal = CStr(vRange.Value)
    '        vRange = CType(vgHoja.Cells(k, 3), Range)
    '        vFecha = Format(CDate(vRange.Value), "yyyy-MM-dd")
    '        vRange = CType(vgHoja.Cells(k, 6), Range)
    '        vCol6 = CStr(vRange.Value)
    '        vRange = CType(vgHoja.Cells(k, 5), Range)
    '        vCol5 = CStr(vRange.Value)
    '        vRange = CType(vgHoja.Cells(k, 4), Range)
    '        vCol4 = CStr(vRange.Value)
    '        vRange = CType(vgHoja.Cells(k, 2), Range)
    '        vCol2 = CStr(vRange.Value)

    '        Select Case vCol2
    '            Case "sms_res_enviadas"
    '                PProcesaRegistroPc01(vShortcode, vTotal, vFecha, vCol4, vCol5, vCol6)
    '            Case "sms_enviados"
    '                PProcesaRegistroPc02(vShortcode, vTotal, vFecha, vCol4, vCol5, vCol6)
    '            Case "sms_enviados_ol"
    '                PProcesaRegistroPc03(vShortcode, vTotal, vFecha, vCol4, vCol5, vCol6)
    '        End Select

    '        FPActualizaProgreso(k, vCantFilasXls, vgPorcionLecturaReg, ENTipoAcumulador.ProcesoLectura)
    '        k = k + 1
    '        vRange = CType(vgHoja.Cells(k, 1), Range)
    '    End While
    'End Sub
    Private Sub PLecturaDatosPcOldb()
        Dim k As Integer
        Dim vShortcode As String
        Dim vFecha As String
        Dim vTotal As String
        Dim vCol2 As String
        Dim vCol4 As String
        Dim vCol6 As String
        Dim vCol5 As String
        Dim vReg As String
        Dim vCantFilasXls As Integer
        Dim vDtDatosOrigen As New DataTable
        Dim vLectorArchivo As StreamReader = Nothing
        Dim vLenStr As Integer
        Dim vPosStr As Integer
        Dim vParametros As New ParametrosCmd
        Dim vDicCols As New Dictionary(Of String, String)
        Dim vDatos As Object = Nothing
        Try
            'vDtDatosOrigen = FObtenerDataOrigen()


            vLectorArchivo = New StreamReader(vgProcArchivo)
            'vCantFilasXls = vDtDatosOrigen.Rows.Count
            vCantFilasXls = Utilidades.CtrlArchivos.FContarRegistrosArchivo(vgProcArchivo) - 4
            'Nota: se resta 4 porque en la cabecera se tienen 2 lineas y al final del archivo hay 2 lineas que no representan información del banco

            'Nos saltamos dos lineas porque son datos de la cabecera
            Try
                vLectorArchivo.ReadLine()
                vLectorArchivo.ReadLine()
            Catch ex As Exception
                Dim exGen01 As New Exception("El archivo no tiene un formato correcto")
                Throw exGen01
            End Try

            'Obtengo el tamaño de las columnas
            With vParametros
                .Campos.Append("ccb_identificacion, ccb_valor")
                .Condicion.Append("ccb_emisor = 'PACIFICARD'")
                .Orden.Append("ccb_identificacion")
                .Tabla = "sms_config_carga_bancos"
            End With

            If vDbControlPlataforma.FdbSelectSql("msgswitchweb", vParametros, vDatos) Then
                Dim vConfigColumnas = DirectCast(vDatos, List(Of List(Of Object)))
                For Each vItem As List(Of Object) In vConfigColumnas
                    vDicCols.Add(vItem(0).ToString, vItem(1).ToString)
                Next
            End If
            Do
                vPosStr = 0
                vReg = vLectorArchivo.ReadLine()
                If vReg IsNot Nothing AndAlso vReg <> "" Then
                    'Shortcode
                    vLenStr = CInt(vDicCols("es_shortcode"))
                    vShortcode = vReg.Substring(0, vLenStr).Trim
                    'es_tabla
                    vPosStr = vPosStr + vLenStr
                    vLenStr = CInt(vDicCols("es_tabla"))
                    vCol2 = vReg.Substring(vPosStr, vLenStr).Trim
                    'es_fecha
                    vPosStr = vPosStr + vLenStr
                    vLenStr = CInt(vDicCols("es_fecha"))
                    vFecha = vReg.Substring(vPosStr, vLenStr).Trim.Substring(0, 10)
                    'es_tipo
                    vPosStr = vPosStr + vLenStr
                    vLenStr = CInt(vDicCols("es_tipo"))
                    vCol4 = vReg.Substring(vPosStr, vLenStr).Trim
                    'es_codigo
                    vPosStr = vPosStr + vLenStr
                    vLenStr = CInt(vDicCols("es_codigo"))
                    vCol5 = vReg.Substring(vPosStr, vLenStr).Trim
                    'es_descripcion
                    vPosStr = vPosStr + vLenStr
                    vLenStr = CInt(vDicCols("es_descripcion"))
                    vCol6 = vReg.Substring(vPosStr, vLenStr).Trim
                    'es_valor
                    vPosStr = vPosStr + vLenStr
                    vLenStr = CInt(vDicCols("es_valor"))
                    vTotal = vReg.Substring(vPosStr, vLenStr).Trim

                    Select Case vCol2
                        Case "sms_res_enviadas"
                            PProcesaRegistroPc01(vShortcode, vTotal, vFecha, vCol4, vCol5, vCol6)
                        Case "sms_enviados"
                            PProcesaRegistroPc02(vShortcode, vTotal, vFecha, vCol4, vCol5, vCol6)
                        Case "sms_enviados_ol"
                            PProcesaRegistroPc03(vShortcode, vTotal, vFecha, vCol4, vCol5, vCol6)
                    End Select

                    FPActualizaProgreso(k, vCantFilasXls, vgPorcionLecturaReg, ENTipoAcumulador.ProcesoLectura)
                    k = k + 1
                End If
            Loop Until vReg Is Nothing OrElse vReg = ""

        Catch ex As Exception
            Dim exGen As New Exception(ex.Message)
            Throw exGen
        Finally
            vDtDatosOrigen = Nothing
        End Try
    End Sub
    Private Function FObtenerValor(ByVal parDato As Object) As Object
        Dim vResultado As Object
        If parDato IsNot Nothing AndAlso Not IsDBNull(parDato) Then
            vResultado = parDato
        Else
            vResultado = ""
        End If
        Return vResultado
    End Function
    Private Function FObtieneCmd(ByVal parHojaXls As String) As String
        Dim vSentencia As String = "SELECT"
        Dim vTablaOrigen As String = parHojaXls

        Return String.Format("{0} {1} FROM {2}", vSentencia, "*", "[" & vTablaOrigen & "]")
    End Function
    Private Function FObtenerDataOrigen() As DataTable
        Dim vResult As New DataTable
        Dim vStringCnx As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & vgProcArchivo & ";Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1';"
        oCnn = New OleDbConnection(vStringCnx)
        Try
            oCnn.Open()
            Using vCmd As New OleDbCommand
                vCmd.Connection = oCnn
                vCmd.CommandType = CommandType.Text
                vCmd.CommandText = FObtieneCmd(vgProcXlsHoja)
                Using vOleda As New OleDbDataAdapter(vCmd)
                    vOleda.Fill(vResult)
                End Using
            End Using
            Return vResult
        Catch ex As Exception
            Dim exGen As New Exception("No se pudo leer el archivo Excel. " & ex.Message)
            Throw exGen
        Finally
            If oCnn.State <> ConnectionState.Closed Then
                oCnn.Close()
            End If
            oCnn = Nothing
        End Try
    End Function
    'Private Sub PLecturaDatosBg()
    '    Dim k As Integer
    '    Dim vShortcode As String
    '    Dim vFecha As String
    '    Dim vTotal As String
    '    Dim vCol2 As String
    '    Dim vCol4 As String
    '    Dim vCol6 As String
    '    Dim vCol5 As String
    '    Dim vCantFilasXls As Integer
    '    Dim vRange As Range

    '    vgHoja = CType(vgLibro.Worksheets(vgProcXlsHoja), Excel.Worksheet)
    '    vgHoja.Activate()
    '    vCantFilasXls = vgHoja.Range("a1").CurrentRegion.Rows.Count
    '    k = 2
    '    vRange = CType(vgHoja.Cells(k, 1), Range)
    '    While Not String.IsNullOrEmpty(Trim(CStr(vRange.Value)))
    '        vRange = CType(vgHoja.Cells(k, 1), Range)
    '        vShortcode = CStr(vRange.Value)
    '        vRange = CType(vgHoja.Cells(k, 7), Range)
    '        vTotal = CStr(vRange.Value)
    '        vRange = CType(vgHoja.Cells(k, 3), Range)
    '        vFecha = Format(CDate(vRange.Value), "yyyy-MM-dd")
    '        vRange = CType(vgHoja.Cells(k, 6), Range)
    '        vCol6 = CStr(vRange.Value)
    '        vRange = CType(vgHoja.Cells(k, 5), Range)
    '        vCol5 = CStr(vRange.Value)
    '        vRange = CType(vgHoja.Cells(k, 4), Range)
    '        vCol4 = CStr(vRange.Value)
    '        vRange = CType(vgHoja.Cells(k, 2), Range)
    '        vCol2 = CStr(vRange.Value)

    '        If vCol4 = "P" OrElse vCol4 = "M" OrElse vCol4 = "OPE" OrElse vCol4 = "A" Then
    '            Select Case vCol2
    '                Case "sms_res_enviadas"
    '                    PProcesaRegistroBg01(vShortcode, vTotal, vFecha, vCol5, vCol6)
    '                Case "sms_enviados"
    '                    PProcesaRegistroBg02(vShortcode, vTotal, vFecha, vCol4, vCol5, vCol6)
    '                Case "sms_enviados_ol"
    '                    PProcesaRegistroBg03(vShortcode, vTotal, vFecha, vCol4, vCol5, vCol6)
    '            End Select
    '        End If
    '        FPActualizaProgreso(k, vCantFilasXls, vgPorcionLecturaReg, ENTipoAcumulador.ProcesoLectura)
    '        k = k + 1
    '    End While
    'End Sub
    Private Sub PLecturaDatosBgOledb()
        Dim k As Integer
        Dim vShortcode As String
        Dim vFecha As String
        Dim vTotal As String
        Dim vCol2 As String
        Dim vCol4 As String
        Dim vCol6 As String
        Dim vCol5 As String
        Dim vCantFilasXls As Integer
        Dim vDtDatosOrigen As New DataTable
        Try
            vDtDatosOrigen = FObtenerDataOrigen()

            vCantFilasXls = vDtDatosOrigen.Rows.Count
            k = 2 'La lectura inicia desde la segunda fila
            If vDtDatosOrigen.Rows.Count >= 0 Then
                Do
                    vShortcode = CStr(FObtenerValor(vDtDatosOrigen.Rows(k - 1)(0)))
                    vTotal = CStr(FObtenerValor(vDtDatosOrigen.Rows(k - 1)(6)))
                    vFecha = Format(CDate(FObtenerValor(vDtDatosOrigen.Rows(k - 1)(2))), "yyyy-MM-dd")
                    vCol6 = CStr(FObtenerValor(vDtDatosOrigen.Rows(k - 1)(5)))
                    vCol5 = CStr(FObtenerValor(vDtDatosOrigen.Rows(k - 1)(4)))
                    vCol4 = CStr(FObtenerValor(vDtDatosOrigen.Rows(k - 1)(3)))
                    vCol2 = CStr(FObtenerValor(vDtDatosOrigen.Rows(k - 1)(1)))

                    If vCol4 = "P" OrElse vCol4 = "M" OrElse vCol4 = "OPE" OrElse vCol4 = "A" Then
                        Select Case vCol2
                            Case "sms_res_enviadas"
                                PProcesaRegistroBg01(vShortcode, vTotal, vFecha, vCol5, vCol6)
                            Case "sms_enviados"
                                PProcesaRegistroBg02(vShortcode, vTotal, vFecha, vCol4, vCol5, vCol6)
                            Case "sms_enviados_ol"
                                PProcesaRegistroBg03(vShortcode, vTotal, vFecha, vCol4, vCol5, vCol6)
                        End Select
                    End If
                    FPActualizaProgreso(k, vCantFilasXls, vgPorcionLecturaReg, ENTipoAcumulador.ProcesoLectura)
                    k = k + 1
                Loop Until k > vDtDatosOrigen.Rows.Count
            End If
        Catch ex As Exception
            Dim exGen As New Exception(ex.Message)
            Throw exGen
        Finally
            vDtDatosOrigen = Nothing
        End Try
    End Sub
    Private Function FObtenerTablaTotales(ByVal parTipoPeriodo As String, ByVal parConReplicacion As String) As String
        If parTipoPeriodo = "Mensual" AndAlso parConReplicacion = "No" Then
            Return "sms_totales_bancos_tmp"
        Else
            Return "sms_totales_bancos"
        End If
    End Function
    Private Sub PCargarDatosBancosBG()

        Dim vTablaTotales As String = FObtenerTablaTotales(vgProcTipo, vgProcReplicar)

        Try
            PAbrirConexionesParaCargarDatos()

            'PInicializaExcel()

            'Si se ha seleccionado hoja del archivo excel
            If vgProcXlsHoja <> "" AndAlso vgProcXlsHoja <> "NO APLICA" Then

                PLecturaDatosBgOledb()

                If vgTotBancosds.Tables(0).Rows.Count > 0 Then
                    vDbControlPlataforma.FdbInsertBulkSql("msgswitchweb", vTablaTotales, vgTotBancosds.Tables(0))
                End If
                FPActualizaProgreso(2, vgPorcionGrabarData, vgPorcionGrabarData, ENTipoAcumulador.ProcesoGrabar)
                If vTablaTotales <> "sms_totales_bancos_tmp" Then
                    If vgTotIntermBancosds.Tables(0).Rows.Count > 0 Then
                        vDbControlPlataforma.FdbInsertBulkSql("msgswitchweb", "sms_diarios_bancos", vgTotIntermBancosds.Tables(0))
                    End If

                    FPActualizaProgreso(3, vgPorcionGrabarData, vgPorcionGrabarData, ENTipoAcumulador.ProcesoGrabar)

                    'Realizo la replicación
                    If vgProcReplicar = "Si" AndAlso vgTotBancosds.Tables(0).Rows.Count > 0 Then
                        vDbControlReporteria.FdbInsertBulkSql("sms_reportes", "sms_totales_bancos", vgTotBancosds.Tables(0))
                    End If
                    FPActualizaProgreso(4, vgPorcionGrabarData, vgPorcionGrabarData, ENTipoAcumulador.ProcesoGrabar)
                End If
            End If
        Catch ex As Exception
            PConsola("Error al cargar la información: " & ex.Message)
            Dim exGen As New Exception("Error al cargar la información: " & ex.Message)
            Throw exGen
        Finally
            vDbControlPlataforma.FdbCierreConexionSql()
            vDbControlReporteria.FdbCierreConexionSql()
        End Try
    End Sub
    Private Function FObtenerOperadoraBg(ByVal parDato As String) As String
        Dim vResultado As String = ""
        Dim vCadena As String()
        vCadena = Split(parDato, "-")
        For x As Integer = 0 To vCadena.Count
            If vCadena(x).Length = 1 Then
                vResultado = vCadena(x)
                x = vCadena.Count + 1 'Truncamos el for (reemplazo 'exit for')
            End If
        Next
        Return vResultado
    End Function
    Private Sub PProcesaRegistroBg03(ByVal parShortcode As String,
                                     ByVal parTotal As String,
                                     ByVal parFecha As String,
                                     ByVal parCol4 As String,
                                     ByVal parCol5 As String,
                                     ByVal parCol6 As String)
        'sms_enviados_ol
        Dim vTipo As String = ""
        Dim vOperadora As String = ""
        If parCol4 = "P" Then
            vTipo = "E"
            vOperadora = Mid(parCol5, 1, 1)
        ElseIf parCol4 = "OPE" AndAlso Microsoft.VisualBasic.Left(parCol5, 5) = "BGMSG" Then
            vTipo = "OL"
            vOperadora = parCol5
        ElseIf parCol4 = "OPE" Then
            Dim vTipoCod As String = Microsoft.VisualBasic.Left(parCol5, 5)
            If vgItmsTipoCods.Contains(vTipoCod) Then
                vTipo = "OL"
                vOperadora = parCol5
            End If
        End If

        If vTipo <> "" Then
            'Si vTipo viene lleno, entonces se procede a insertar...
            If parCol6 = "ENVIADOS" Then
                PLLenaTablaTotales(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, parTotal)
            ElseIf Not vgItmsOperadoras.Contains(parCol6) Then
                PLlenaTotalesIntermedios(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, CInt(parTotal), parCol6)
            End If
        End If

    End Sub
    Private Sub PProcesaRegistroBg02(ByVal parShortcode As String,
                                     ByVal parTotal As String,
                                     ByVal parFecha As String,
                                     ByVal parCol4 As String,
                                     ByVal parCol5 As String,
                                     ByVal parCol6 As String)
        'sms_enviados
        Dim vTipo As String = ""
        Dim vOperadora As String = ""
        Select Case parCol4
            Case "P"
                vTipo = "E"
                vOperadora = Mid(parCol5, 1, 1)
            Case "A"
                vTipo = "A"
                vOperadora = Mid(parCol5, 1, 1)
            Case "M"
                vTipo = "C"
                vOperadora = Mid(parCol5, 1, 1)
            Case "OPE"
                If Microsoft.VisualBasic.Left(parCol5, 3) = "WEB" Then
                    vTipo = "WEB"
                    vOperadora = FObtenerOperadoraBg(parCol5)
                ElseIf Microsoft.VisualBasic.Left(parCol5, 5) = "LGSEG" Then
                    vTipo = "LGSEG"
                    If parCol6 = "ENVIADOS" Then
                        vOperadora = FObtenerOperadoraBg(parCol5)
                    ElseIf Not vgItmsOperadoras.Contains(parCol6) Then
                        vOperadora = Mid(parCol5, 7, InStr(parCol5, "-ERR") - 7)
                    End If
                ElseIf Microsoft.VisualBasic.Left(parCol5, 2) = "C-" Then
                    vTipo = "Ñ"
                    vOperadora = FObtenerOperadoraBg(parCol5)
                End If

        End Select



        'If parCol4 = "P" Then
        '    vTipo = "E"
        '    vOperadora = Mid(parCol5, 1, 1)
        'ElseIf parCol4 = "A" Then
        '    vTipo = "A"
        '    vOperadora = Mid(parCol5, 1, 1)
        'ElseIf parCol4 = "M" Then
        '    vTipo = "C"
        '    vOperadora = Mid(parCol5, 1, 1)
        'ElseIf parCol4 = "OPE" AndAlso Microsoft.VisualBasic.Left(parCol5, 3) = "WEB" Then
        '    vTipo = "WEB"
        '    vOperadora = FObtenerOperadoraBg(parCol5)
        'ElseIf parCol4 = "OPE" AndAlso Microsoft.VisualBasic.Left(parCol5, 5) = "LGSEG" Then
        '    vTipo = "LGSEG"
        '    If parCol6 = "ENVIADOS" Then
        '        vOperadora = FObtenerOperadoraBg(parCol5)
        '    ElseIf Not vgItmsOperadoras.Contains(parCol6) Then
        '        vOperadora = Mid(parCol5, 7, InStr(parCol5, "-ERR") - 7)
        '    End If
        'ElseIf parCol4 = "OPE" AndAlso Microsoft.VisualBasic.Left(parCol5, 2) = "C-" Then
        '    vTipo = "Ñ"
        '    vOperadora = FObtenerOperadoraBg(parCol5)
        'End If

        If vTipo <> "" Then
            'Si vTipo viene lleno, entonces se procede a insertar...
            If parCol6 = "ENVIADOS" Then
                PLLenaTablaTotales(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, parTotal)
            ElseIf Not vgItmsOperadoras.Contains(parCol6) Then
                PLlenaTotalesIntermedios(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, CInt(parTotal), parCol6)
            End If
        End If
    End Sub
    Private Sub PProcesaRegistroBg01(ByVal parShortcode As String,
                                     ByVal parTotal As String,
                                     ByVal parFecha As String,
                                     ByVal parCol5 As String,
                                     ByVal parCol6 As String)
        'sms_res_enviadas
        Dim vTipo As String = "R"
        Dim vOperadora As String
        vOperadora = Mid(parCol5, 1, 1)
        If parCol6 = "ENVIADOS" Then
            PLLenaTablaTotales(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, parTotal)
        ElseIf Not vgItmsOperadoras.Contains(parCol6) Then
            PLlenaTotalesIntermedios(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, CInt(parTotal), parCol6)
        End If
    End Sub
    Private Sub PLLenaTablaTotales(ByVal parEmisor As String, ByVal parShortcode As String, ByVal parTipo As String, ByVal parOperadora As String, ByVal parFecha As String, ByVal parTotal As String)
        Try
            Dim vClavePrim(4) As Object
            Dim vTipoTmp As String = ""
            Dim vRowNueva As DataRow = vgTotBancosds.Tables(0).NewRow
            Dim vRowFiltrado As DataRow = Nothing

            If vgProcTipo = "Diario" AndAlso parTipo = "WEB" Then
                vTipoTmp = "A"
                vClavePrim(0) = parEmisor
                vClavePrim(1) = parShortcode
                vClavePrim(2) = vTipoTmp
                vClavePrim(3) = parOperadora
                vClavePrim(4) = parFecha
                Try
                    vRowFiltrado = vgTotBancosds.Tables(0).Rows.Find(vClavePrim)
                Catch ex As Exception
                    vRowFiltrado = Nothing
                End Try
                If vRowFiltrado IsNot Nothing Then
                    vRowFiltrado("Total") = CInt(vRowFiltrado("Total")) - CInt(parTotal)
                    vgRegCargados = vgRegCargados + 1
                End If
            End If

            vRowFiltrado = Nothing
            vClavePrim(0) = parEmisor
            vClavePrim(1) = parShortcode
            vClavePrim(2) = parTipo
            vClavePrim(3) = parOperadora
            vClavePrim(4) = parFecha
            vRowFiltrado = vgTotBancosds.Tables(0).Rows.Find(vClavePrim)

            If vRowFiltrado IsNot Nothing Then
                vRowFiltrado("Total") = CInt(vRowFiltrado("Total")) + CInt(parTotal)
                vgRegCargados = vgRegCargados + 1
            Else
                vRowNueva("Emisor") = parEmisor
                vRowNueva("Shortcode") = parShortcode
                vRowNueva("Tipo") = parTipo
                vRowNueva("Operadora") = parOperadora
                vRowNueva("Fecha") = parFecha
                vRowNueva("Total") = parTotal
                vgTotBancosds.Tables(0).Rows.Add(vRowNueva)
                vgRegCargados = vgRegCargados + 1
            End If
        Catch ex As Exception
            vgRegError = vgRegError + 1
        End Try
    End Sub
    Private Sub PLlenaTotalesIntermedios(ByVal parEmisor As String,
                                         ByVal parShortcode As String,
                                         ByVal parTipo As String,
                                         ByVal parOperadora As String,
                                         ByVal parFecha As String,
                                         ByVal parTotal As Integer,
                                         ByVal parDescripcion As String)
        Dim vRowNueva As DataRow = vgTotIntermBancosds.Tables(0).NewRow
        Dim vRowFiltrado As DataRow = Nothing
        Dim vClavePrim(5) As Object
        Try
            vRowFiltrado = Nothing
            vClavePrim(0) = parEmisor
            vClavePrim(1) = parShortcode
            vClavePrim(2) = parTipo
            vClavePrim(3) = parOperadora
            vClavePrim(4) = parFecha
            vClavePrim(5) = parDescripcion
            Try
                vRowFiltrado = vgTotIntermBancosds.Tables(0).Rows.Find(vClavePrim)
            Catch ex As Exception
                vRowFiltrado = Nothing
            End Try


            If vRowFiltrado IsNot Nothing Then
                vRowFiltrado("Total") = CInt(vRowFiltrado("Total")) + CInt(parTotal)
                vgRegCargados = vgRegCargados + 1
            Else
                vRowNueva("Emisor") = parEmisor
                vRowNueva("Shortcode") = parShortcode
                vRowNueva("Tipo") = parTipo
                vRowNueva("Operadora") = parOperadora
                vRowNueva("Fecha") = parFecha
                vRowNueva("Total") = parTotal
                vRowNueva("Estado") = parDescripcion
                vgTotIntermBancosds.Tables(0).Rows.Add(vRowNueva)
                vgRegCargados = vgRegCargados + 1
            End If
        Catch ex As Exception
            vgRegError = vgRegError + 1
        End Try
    End Sub
    Private Sub PAbrirConexionesParaCargarDatos()
        Try
            If Not vDbControlPlataforma.FdbConexionSql(My.Settings.vgServerPlataforma) Then
                Dim exPlat As New Exception("No se pudo establecer la conexión con la base de datos de la plataforma transaccional.")
                Throw exPlat
            End If
            If Not vDbControlReporteria.FdbConexionSql(My.Settings.vgServerReporteria) Then
                Dim exRep As New Exception("No se pudo establecer la conexión con la base de datos de reportería.")
                Throw exRep
            End If
        Catch ex As Exception
            PConsola(ex.Message)
            vDbControlPlataforma.FdbCierreConexionSql()
            vDbControlReporteria.FdbCierreConexionSql()
            Dim exGen As New Exception(ex.Message)
            Throw exGen
        End Try


    End Sub

    Private Sub PCargarDatosBancos()
        vgRegCargados = 0
        vgRegError = 0
        Dim vFileSystem As StreamReader = Nothing
        Dim vDelimitador As String
        Dim vReg As String
        Dim vCampos As String()
        Dim k As Integer = 1
        Dim vShortcode As String
        Dim vTipo As String
        Dim vOperadora As String
        Dim vFecha As String
        Dim vTotal As Integer
        Dim vTablaTotales As String = "sms_totales_bancos"
        Dim vCantFilasTxt As Integer = 0
        Try

            PAbrirConexionesParaCargarDatos()
            Try
                vFileSystem = New StreamReader(vgProcArchivo)
            Catch ex As Exception
                Dim exFile As New Exception("No se pudo acceder al archivo. Ruta: " & vgProcArchivo)
                PConsola("No se pudo acceder al archivo. Ruta: " & vgProcArchivo)
                Throw exFile
            End Try

            vCantFilasTxt = FContarRegistrosArchivo(vgProcArchivo)

            vDelimitador = Chr(9)
            Do
                vReg = vFileSystem.ReadLine()

                If vReg IsNot Nothing AndAlso Not String.IsNullOrEmpty(vReg) Then
                    vCampos = Split(vReg, vDelimitador)
                    vShortcode = vCampos(0)
                    vTipo = vCampos(1)
                    vOperadora = vCampos(2)
                    vFecha = Format(CDate(vCampos(3)), "yyyy-MM-dd")
                    vTotal = CInt(vCampos(4))
                    PLLenaTablaTotales(vgProcEmisor, vShortcode, vTipo, vOperadora, vFecha, CStr(vTotal))
                End If
                FPActualizaProgreso(k, vCantFilasTxt, vgPorcionLecturaReg, ENTipoAcumulador.ProcesoLectura)
                k = k + 1
            Loop Until vReg Is Nothing

            If vgProcTipo = "Mensual" AndAlso vgProcReplicar = "No" Then
                vTablaTotales = "sms_totales_bancos_tmp"
            End If

            If vgTotBancosds.Tables(0).Rows.Count > 0 Then
                vDbControlPlataforma.FdbInsertBulkSql("msgswitchweb", vTablaTotales, vgTotBancosds.Tables(0))
            End If

            FPActualizaProgreso(2, vgPorcionGrabarData, vgPorcionGrabarData, ENTipoAcumulador.ProcesoGrabar)

            If vTablaTotales <> "sms_totales_bancos_tmp" Then
                'En caso de que la información anterior no haya sido grabada en <<sms_totales_bancos>>, entonces no debe grabarse lo demás
                If vgTotIntermBancosds.Tables(0).Rows.Count > 0 Then
                    vDbControlPlataforma.FdbInsertBulkSql("msgswitchweb", "sms_diarios_bancos", vgTotIntermBancosds.Tables(0))
                End If

                FPActualizaProgreso(3, vgPorcionGrabarData, vgPorcionGrabarData, ENTipoAcumulador.ProcesoGrabar)

                'Realizo la replicación
                If vgProcReplicar = "Si" AndAlso vgTotBancosds.Tables(0).Rows.Count > 0 Then
                    vDbControlReporteria.FdbInsertBulkSql("sms_reportes", "sms_totales_bancos", vgTotBancosds.Tables(0))
                End If
                FPActualizaProgreso(4, vgPorcionGrabarData, vgPorcionGrabarData, enTipoAcumulador.ProcesoGrabar)
            End If
        Catch ex As Exception
            PConsola("Error al cargar la información: " & ex.Message)
            Dim exGen As New Exception("Error al cargar la información: " & ex.Message)
            Throw exGen
        Finally
            vDbControlPlataforma.FdbCierreConexionSql()
            vDbControlReporteria.FdbCierreConexionSql()
        End Try

    End Sub
    Private Sub PProcesaRegistroPc01(ByVal parShortcode As String,
                                     ByVal parTotal As String,
                                     ByVal parFecha As String,
                                     ByVal parCol4 As String,
                                     ByVal parCol5 As String,
                                     ByVal parCol6 As String)
        'sms_res_enviadas
        Dim vTipo As String = "R"
        Dim vOperadora As String
        If parCol4 = "P" Then
            vTipo = "R"
            vOperadora = Mid(parCol5, 1, 1)
            If parCol6 = "ENVIADOS" Then
                PLLenaTablaTotales(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, parTotal)
            ElseIf Not vgItmsOperadorasPc.Contains(parCol6) Then
                PLlenaTotalesIntermedios(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, CInt(parTotal), parCol6)
            End If
        End If
    End Sub
    Private Sub PProcesaRegistroPc02(ByVal parShortcode As String,
                                     ByVal parTotal As String,
                                     ByVal parFecha As String,
                                     ByVal parCol4 As String,
                                     ByVal parCol5 As String,
                                     ByVal parCol6 As String)
        'sms_enviados
        Dim vTipo As String = ""
        Dim vOperadora As String = ""
        If parCol4 = "P" Then
            vTipo = "E"
            vOperadora = Mid(parCol5, 1, 1)
        ElseIf parCol4 = "N" AndAlso Microsoft.VisualBasic.Right(parCol5, 1) = "X" Then
            vTipo = "X"
            vOperadora = Mid(parCol5, 1, InStr(parCol5, "-X") - 1)
        ElseIf parCol4 = "N" AndAlso Microsoft.VisualBasic.Right(parCol5, 1) = "N" Then
            vTipo = "N"
            vOperadora = Mid(parCol5, 1, InStr(parCol5, "-N") - 1)
        ElseIf parCol4 = "SO" Then
            vTipo = "SOE"
            vOperadora = parCol5
        ElseIf parCol4 = "OPE" Then
            vTipo = parCol4
            vOperadora = parCol5
        End If
        If vTipo <> "" Then
            If parCol6 = "ENVIADOS" Then
                PLLenaTablaTotales(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, parTotal)
            ElseIf Not vgItmsOperadorasPc.Contains(parCol6) Then
                PLlenaTotalesIntermedios(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, CInt(parTotal), parCol6)
            End If
        End If
    End Sub
    Private Sub PProcesaRegistroPc03(ByVal parShortcode As String,
                                     ByVal parTotal As String,
                                     ByVal parFecha As String,
                                     ByVal parCol4 As String,
                                     ByVal parCol5 As String,
                                     ByVal parCol6 As String)
        'sms_enviados_ol
        Dim vTipo As String = ""
        Dim vOperadora As String = ""
        If parCol4 = "P" Then
            vTipo = "E"
            vOperadora = Mid(parCol5, 1, 1)
        ElseIf parCol4 = "N" AndAlso Microsoft.VisualBasic.Right(parCol5, 1) = "X" Then
            vTipo = "X"
            vOperadora = Mid(parCol5, 1, InStr(parCol5, "-X") - 1)
        ElseIf parCol4 = "N" AndAlso Microsoft.VisualBasic.Right(parCol5, 1) = "N" Then
            vTipo = "N"
            vOperadora = Mid(parCol5, 1, InStr(parCol5, "-N") - 1)
        ElseIf parCol4 = "SO" Then
            vTipo = "SOE"
            vOperadora = parCol5
        ElseIf parCol4 = "OPE" Then
            vTipo = parCol4
            vOperadora = parCol5
        End If
        If vTipo <> "" Then
            If parCol6 = "ENVIADOS" Then
                PLLenaTablaTotales(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, parTotal)
            ElseIf Not vgItmsOperadorasPc.Contains(parCol6) Then
                PLlenaTotalesIntermedios(vgProcEmisor, parShortcode, vTipo, vOperadora, parFecha, CInt(parTotal), parCol6)
            End If
        End If
    End Sub
    Private Sub PCargarDatosBancosPacf()
        Dim vTablaTotales As String = FObtenerTablaTotales(vgProcTipo, vgProcReplicar)
        Try

            PAbrirConexionesParaCargarDatos()

            'Si se ha seleccionado hoja del archivo excel
            'If vgProcXlsHoja <> "" AndAlso vgProcXlsHoja <> "NO APLICA" Then

            PLecturaDatosPcOldb()

            If vgTotBancosds.Tables(0).Rows.Count > 0 Then
                vDbControlPlataforma.FdbInsertBulkSql("msgswitchweb", vTablaTotales, vgTotBancosds.Tables(0))
            End If
            FPActualizaProgreso(2, vgPorcionGrabarData, vgPorcionGrabarData, ENTipoAcumulador.ProcesoGrabar)
            If vTablaTotales <> "sms_totales_bancos_tmp" Then
                If vgTotIntermBancosds.Tables(0).Rows.Count > 0 Then
                    vDbControlPlataforma.FdbInsertBulkSql("msgswitchweb", "sms_diarios_bancos", vgTotIntermBancosds.Tables(0))
                End If

                FPActualizaProgreso(3, vgPorcionGrabarData, vgPorcionGrabarData, ENTipoAcumulador.ProcesoGrabar)

                'Realizo la replicación
                If vgProcReplicar = "Si" AndAlso vgTotBancosds.Tables(0).Rows.Count > 0 Then
                    vDbControlReporteria.FdbInsertBulkSql("sms_reportes", "sms_totales_bancos", vgTotBancosds.Tables(0))
                End If
                FPActualizaProgreso(4, vgPorcionGrabarData, vgPorcionGrabarData, ENTipoAcumulador.ProcesoGrabar)
            End If

            'End If
        Catch ex As Exception
            PConsola("Error al cargar la información: " & ex.Message)
            Dim exGen As New Exception("Error al cargar la información: " & ex.Message)
            Throw exGen
        Finally
            vDbControlPlataforma.FdbCierreConexionSql()
            vDbControlReporteria.FdbCierreConexionSql()
        End Try
    End Sub
    Private Sub PObtieneFecha1()
        Dim vFechaAnt As Date
        Dim vAnio, vMes, vDiaIni, vDiaFin As String

        vFechaAnt = New Date(Now.Year, Now.Month, Now.Day, 0, 0, 0, DateTimeKind.Local)
        vAnio = CStr(vFechaAnt.Year)
        vMes = CStr(vFechaAnt.Month)
        vDiaFin = CStr(Date.DaysInMonth(CInt(vAnio), CInt(vMes)))

        If (vFechaAnt.Month) >= 10 Then
            vMes = CStr(vFechaAnt.Month)
        Else
            vMes = "0" & CStr(vFechaAnt.Month)
        End If

        vDiaIni = "01"

        vgFechaIni = vAnio & "-" & vMes & "-" & vDiaIni
        vgFechaFin = vAnio & "-" & vMes & "-" & vDiaFin
    End Sub
    Private Class CLAcumulador
        Public Property propAvanceEliminacion As Integer = 0
        Public Property propAvanceLectura As Integer = 0
        Public Property propAvanceGrabarData As Integer = 0
    End Class
    Private Enum ENTipoAcumulador
        ProcesoEliminacion
        ProcesoLectura
        ProcesoGrabar
    End Enum
End Module
