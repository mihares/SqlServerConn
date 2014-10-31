Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace Datos
    Public Class DatosSQLServer
        Inherits DatosGlobales.DatosGlobales
        Shared ColComandos As New System.Collections.Hashtable()
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Overrides Property CadenaConexion() As String
            Get
                If Me.mCadenaConexion.Length = 0 Then
                    If Me.mBase.Length <> 0 AndAlso Me.mServidor.Length <> 0 Then
                        Dim sCadena As New System.Text.StringBuilder("")
                        sCadena.Append("data source=<SERVIDOR>;")
                        sCadena.Append("initial catalog=<BASE>;password=<PASSWORD>;")
                        sCadena.Append("persist security info=True;")
                        sCadena.Append("user id=<USER>;packet size=4096")
                        sCadena.Replace("<SERVIDOR>", Me.Servidor)
                        sCadena.Replace("<BASE>", Me.Base)
                        sCadena.Replace("<USER>", Me.Ususario)
                        sCadena.Replace("<PASSWORD>", Me.Password)
                        Return sCadena.ToString()
                    Else
                        Dim Ex As New System.Exception("No se puede establecer la cadena de conexión")
                        Throw Ex
                    End If
                End If
                Return Nothing
            End Get
            Set(ByVal value As String)
                Me.mCadenaConexion = value
            End Set
        End Property
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="Com"></param>
        ''' <param name="Args"></param>
        ''' <remarks></remarks>
        Protected Overloads Overrides Sub CargarParametros(ByVal Com As System.Data.IDbCommand, ByVal Args As Object())
            'Dim Limite As Integer = Com.Parameters.Count
            For i As Integer = 1 To Com.Parameters.Count - 1
                Dim P As System.Data.SqlClient.SqlParameter = DirectCast(Com.Parameters(i), System.Data.SqlClient.SqlParameter)
                If i <= Args.Length Then
                    P.Value = Args(i - 1)
                Else
                    P.Value = Nothing
                End If
            Next
        End Sub
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="Com"></param>
        ''' <param name="Parametros"></param>
        ''' <param name="Args"></param>
        ''' <remarks></remarks>
        Protected Overloads Overrides Sub CargarParametros(ByVal Com As System.Data.IDbCommand, ByVal Parametros() As String, ByVal Args As Object())
            Dim band As Boolean = False
            Dim c As Integer = 1

            While c <= Com.Parameters.Count - 1
                Dim P As System.Data.SqlClient.SqlParameter = DirectCast(Com.Parameters(c), System.Data.SqlClient.SqlParameter)
                For i As Integer = 0 To Parametros.Length - 1
                    If P.ParameterName = Parametros(i) Then
                        band = True
                        P.Value = (Args.GetValue(i))
                        'c = c + 1
                        Exit For
                    End If
                Next
                If band Then
                    band = False
                Else
                    P.Value = Nothing
                    'Com.Parameters.Remove(P)
                End If
                c = c + 1
            End While
        End Sub
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="ProcedimientoAlmacenado"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overloads Overrides Function Comando(ByVal ProcedimientoAlmacenado As String) As System.Data.IDbCommand

            Dim Com As System.Data.SqlClient.SqlCommand
            If ColComandos.Contains(ProcedimientoAlmacenado) Then
                Com = DirectCast(ColComandos(ProcedimientoAlmacenado), System.Data.SqlClient.SqlCommand)
            Else
                Dim Con2 As New System.Data.SqlClient.SqlConnection(Me.CadenaConexion)
                Con2.Open()
                Com = New System.Data.SqlClient.SqlCommand(ProcedimientoAlmacenado, Con2)
                Com.CommandType = System.Data.CommandType.StoredProcedure
                System.Data.SqlClient.SqlCommandBuilder.DeriveParameters(Com)
                Con2.Close()
                Con2.Dispose()

                ColComandos.Add(ProcedimientoAlmacenado, Com)
            End If
            Com.Connection = DirectCast(Me.Conexion, System.Data.SqlClient.SqlConnection)
            Com.Transaction = DirectCast(Me.mTransaccion, System.Data.SqlClient.SqlTransaction)
            Return DirectCast(Com, System.Data.IDbCommand)
        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="Consulta"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overloads Overrides Function ComandoConsulta(ByVal Consulta As String) As System.Data.IDbCommand
            Dim Com As New System.Data.SqlClient.SqlCommand
            Com.CommandType = System.Data.CommandType.Text
            Com.CommandText = Consulta
            Com.Connection = DirectCast(Me.Conexion, System.Data.SqlClient.SqlConnection)
            Return DirectCast(Com, System.Data.IDbCommand)
        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="CadenaConexion"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overloads Overrides Function CrearConexion(ByVal CadenaConexion As String) As System.Data.IDbConnection
            Return DirectCast(New System.Data.SqlClient.SqlConnection(CadenaConexion), System.Data.IDbConnection)
        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="ProcedimientoAlmacenado"></param>
        ''' <param name="Args"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overloads Overrides Function CrearDataAdapter(ByVal ProcedimientoAlmacenado As String, ByVal ParamArray Args As Object()) As System.Data.IDataAdapter
            Dim Da As New System.Data.SqlClient.SqlDataAdapter(DirectCast(Comando(ProcedimientoAlmacenado), System.Data.SqlClient.SqlCommand))
            Da.SelectCommand.CommandTimeout = 200
            If Args.Length <> 0 Then
                CargarParametros(Da.SelectCommand, Args)
            End If
            Return DirectCast(Da, System.Data.IDataAdapter)
        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="Consulta"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overloads Overrides Function CrearDataAdapterConsulta(ByVal Consulta As String) As System.Data.IDataAdapter
            Dim Da As New System.Data.SqlClient.SqlDataAdapter(DirectCast(ComandoConsulta(Consulta), System.Data.SqlClient.SqlCommand))
            Da.SelectCommand.CommandTimeout = 200
            Return DirectCast(Da, System.Data.IDataAdapter)
        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="ProcedimientoAlmacenado"></param>
        ''' <param name="Parametros"></param>
        ''' <param name="Args"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overloads Overrides Function CrearDataAdapter(ByVal ProcedimientoAlmacenado As String, ByVal Parametros() As String, ByVal ParamArray Args As Object()) As System.Data.IDataAdapter
            Dim Da As New System.Data.SqlClient.SqlDataAdapter(DirectCast(Comando(ProcedimientoAlmacenado), System.Data.SqlClient.SqlCommand))
            Da.SelectCommand.CommandTimeout = 200
            If Args.Length <> 0 Then
                CargarParametros(Da.SelectCommand, Parametros, Args)
            End If
            Return DirectCast(Da, System.Data.IDataAdapter)
        End Function

        Public Sub New()
            ' 
            ' TODO: agregar aquí la lógica del constructor 
            ' 
            Me.Base = ""
            Me.Servidor = ""
            Me.Ususario = ""
            Me.Password = ""
        End Sub
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="CadenaConexion"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal CadenaConexion As String)
            Me.CadenaConexion = CadenaConexion
        End Sub

        Public Sub New(ByVal Servidor As String, ByVal Base As String)
            Me.Base = Base
            Me.Servidor = Servidor
        End Sub

        Public Sub New(ByVal Servidor As String, ByVal Base As String, ByVal User As String, ByVal Password As String)
            Me.Base = Base
            Me.Servidor = Servidor
            Me.Ususario = User
            Me.Password = Password
        End Sub

    End Class
End Namespace