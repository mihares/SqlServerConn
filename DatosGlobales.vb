Imports System

Namespace DatosGlobales
    ''' <summary> 
    ''' Clase abstracta de acceso a Datos. 
    ''' </summary> 
    Public MustInherit Class DatosGlobales
#Region "Constructores"
        Public Sub New()
        End Sub
#End Region
#Region "Declaración de Atributos"
        Protected mServidor As String = ""
        Protected mBase As String = ""
        Protected mUsuario As String = ""
        Protected mPassword As String = ""
        Protected mCadenaConexion As String = ""
        Protected mConexion As System.Data.IDbConnection
#End Region
#Region "Propiedades"
        ''' <summary> 
        ''' Nombre del equipo servidor de datos. 
        ''' </summary> 
        Public Property Servidor() As String
            Get
                Return mServidor
            End Get
            Set(ByVal value As String)
                mServidor = value
            End Set
        End Property
        ''' <summary> 
        ''' Nombre de la base de datos a utilizar. 
        ''' </summary> 
        Public Property Base() As String
            Get
                Return mBase
            End Get
            Set(ByVal value As String)
                mBase = value
            End Set
        End Property
        ''' <summary> 
        ''' Nombre del Usuario de la BD. 
        ''' </summary> 
        Public Property Ususario() As String
            Get
                Return mUsuario
            End Get
            Set(ByVal value As String)
                mUsuario = value
            End Set
        End Property
        ''' <summary> 
        ''' Password del Usuario de la BD. 
        ''' </summary> 
        Public Property Password() As String
            Get
                Return mPassword
            End Get
            Set(ByVal value As String)
                mPassword = value
            End Set
        End Property

        ''' <summary> 
        ''' Cadena de conexión completa a la base. 
        ''' </summary> 

        Public MustOverride Property CadenaConexion() As String

#End Region
#Region "Privadas"
        ''' <summary> 
        ''' Crea u obtiene un objeto para conectarse a la base de dtaos. 
        ''' </summary> 

        Protected ReadOnly Property Conexion() As System.Data.IDbConnection
            Get
                If mConexion Is Nothing Then
                    mConexion = CrearConexion(Me.CadenaConexion)
                End If
                If mConexion.State <> System.Data.ConnectionState.Open Then
                    mConexion.Open()
                End If
                Return mConexion
            End Get
        End Property
#End Region
#Region "Lecturas"
        ''' <summary> 
        ''' Obtiene un DataSet a partir de un Procedimiento Almacenado. 
        ''' </summary> 
        Public Function TraerDataset(ByVal ProcedimientoAlmacenado As String) As System.Data.DataSet
            Dim mDataSet As New System.Data.DataSet()
            Me.CrearDataAdapter(ProcedimientoAlmacenado).Fill(mDataSet)
            Return mDataSet

        End Function
        ''' <summary> 
        ''' Obtiene un DataSet a partir de un Procedimiento Almacenado y sus parámetros. 
        ''' </summary> 
        Public Function TraerDataset(ByVal ProcedimientoAlmacenado As String, ByVal ParamArray Args As Object()) As System.Data.DataSet
            Dim mDataSet As New System.Data.DataSet()
            Me.CrearDataAdapter(ProcedimientoAlmacenado, Args).Fill(mDataSet)
            Return mDataSet
        End Function
        ''' <summary>
        ''' Obtiene un DataSet a partir de un Procedimiento Almacenado y sus parámetros de entrada y salida
        ''' </summary>
        ''' <param name="ProcedimientoAlmacenado"></param>
        ''' <param name="Parametros"></param>
        ''' <param name="Args"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function TraerDataset(ByVal ProcedimientoAlmacenado As String, ByVal Parametros() As String, ByVal ParamArray Args As Object()) As System.Data.DataSet
            Dim mDataSet As New System.Data.DataSet()
            Me.CrearDataAdapter(ProcedimientoAlmacenado, Parametros, Args).Fill(mDataSet)
            Return mDataSet
        End Function
        ''' <summary>
        ''' Obtiene un DataSet a partir de un Query Sql.
        ''' </summary>
        ''' <param name="Consulta"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function TraerDatasetConsulta(ByVal Consulta As String) As System.Data.DataSet
            Dim mDataSet As New System.Data.DataSet()
            Me.CrearDataAdapterConsulta(Consulta).Fill(mDataSet)
            Return mDataSet
        End Function
        ''' <summary> 
        ''' Obtiene un DataTable a partir de un Procedimiento Almacenado. 
        ''' </summary> 
        Public Function TraerDataTable(ByVal ProcedimientoAlmacenado As String) As System.Data.DataTable
            Return TraerDataset(ProcedimientoAlmacenado).Tables(0).Copy()
        End Function
        ''' <summary>
        ''' btiene un DataTable a partir de un Query Sql.
        ''' </summary>
        ''' <param name="Consulta"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function TraerDataTableConsulta(ByVal Consulta As String) As System.Data.DataTable
            Return TraerDatasetConsulta(Consulta).Tables(0).Copy()
        End Function
        ''' <summary> 
        ''' Obtiene un DataSet a partir de un Procedimiento Almacenado y sus parámetros. 
        ''' </summary> 
        Public Function TraerDataTable(ByVal ProcedimientoAlmacenado As String, ByVal ParamArray Args As Object()) As System.Data.DataTable
            Return TraerDataset(ProcedimientoAlmacenado, Args).Tables(0).Copy()
        End Function
        ''' <summary>
        ''' Obtiene un DataTable a partir de un Procedimiento Almacenado y sus parámetros de entrada y salida
        ''' </summary>
        ''' <param name="ProcedimientoAlmacenado"></param>
        ''' <param name="Parametros"></param>
        ''' <param name="Args"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function TraerDataTable(ByVal ProcedimientoAlmacenado As String, ByVal Parametros() As String, ByVal ParamArray Args As Object()) As System.Data.DataTable
            Return TraerDataset(ProcedimientoAlmacenado, Parametros, Args).Tables(0).Copy()
        End Function
        ''' <summary> 
        ''' Obtiene un DataReader a partir de un Procedimiento Almacenado. 
        ''' </summary> 
        Public Function TraerDataReader(ByVal ProcedimientoAlmacenado As String) As System.Data.IDataReader
            Dim mDataSet As New System.Data.DataSet()
            Me.CrearDataAdapter(ProcedimientoAlmacenado).Fill(mDataSet)
            Return mDataSet.Tables(0).CreateDataReader()
        End Function
        ''' <summary> 
        ''' Obtiene un DataReader a partir de un Procedimiento Almacenado y sus parámetros. 
        ''' </summary> 
        Public Function TraerDataReader(ByVal ProcedimientoAlmacenado As String, ByVal ParamArray Args As Object()) As System.Data.IDataReader
            Dim mDataSet As New System.Data.DataSet()
            Me.CrearDataAdapter(ProcedimientoAlmacenado, Args).Fill(mDataSet)
            Return mDataSet.Tables(0).CreateDataReader()
        End Function
        ''' <summary> 
        ''' Obtiene un Valor a partir de un Procedimiento Almacenado. 
        ''' </summary> 
        Public Function TraerValor(ByVal ProcedimientoAlmacenado As String) As Object
            Dim Com As System.Data.IDbCommand = Comando(ProcedimientoAlmacenado)
            Com.ExecuteNonQuery()
            Dim Resp As Object = Nothing
            For Each Par As System.Data.IDbDataParameter In Com.Parameters
                If Par.Direction = System.Data.ParameterDirection.InputOutput OrElse Par.Direction = System.Data.ParameterDirection.Output Then
                    Resp = Par.Value
                End If
            Next
            Return Resp
        End Function

        ''' <summary> 
        ''' Obtiene un Valor Escalar a partir de un Procedimiento Almacenado. 
        ''' </summary> 
        Public Function TraerValorEscalar(ByVal ProcedimientoAlmacenado As String) As Object
            Dim Com As System.Data.IDbCommand = Comando(ProcedimientoAlmacenado)
            Return Com.ExecuteScalar()
        End Function
        ''' <summary> 
        ''' Obtiene un Valor Escalar a partir de un Procedimiento Almacenado, con Params de Entrada
        ''' </summary> 
        Public Function TraerValorEscalar(ByVal ProcedimientoAlmacenado As String, ByVal ParamArray Args As Object()) As Object
            Dim Com As System.Data.IDbCommand = Comando(ProcedimientoAlmacenado)
            CargarParametros(Com, Args)
            Return Com.ExecuteScalar()
        End Function

        ''' <summary> 
        ''' Obtiene un Valor Escalar a partir de un Procedimiento Almacenado, con Params de Entrada y parametros definidos
        ''' </summary> 
        Public Function TraerValorEscalar(ByVal ProcedimientoAlmacenado As String, ByVal Parametros() As String, ByVal ParamArray Args As Object()) As Object
            Dim Com As System.Data.IDbCommand = Comando(ProcedimientoAlmacenado)
            CargarParametros(Com, Parametros, Args)
            Com.ExecuteScalar()
            'retrono los valores de la consulta del procedimiento tambien con la cantidad de filas afectadas por defecto
            Dim Par As System.Data.IDbDataParameter = DirectCast(Com.Parameters(0), System.Data.IDbDataParameter)
            Return Par.Value
        End Function

        ''' <summary> 
        ''' Obtiene un Valor a partir de un Procedimiento Almacenado, y sus parámetros. 
        ''' </summary> 
        Public Function TraerValor(ByVal ProcedimientoAlmacenado As String, ByVal ParamArray Args As Object()) As Object
            Dim Com As System.Data.IDbCommand = Comando(ProcedimientoAlmacenado)
            CargarParametros(Com, Args)
            Com.ExecuteNonQuery()
            Dim Resp As Object = Nothing
            For Each Par As System.Data.IDbDataParameter In Com.Parameters
                If Par.Direction = System.Data.ParameterDirection.InputOutput OrElse Par.Direction = System.Data.ParameterDirection.Output OrElse Par.Direction = System.Data.ParameterDirection.ReturnValue Then
                    Resp = Par.Value
                End If
            Next
            Return Resp
        End Function
#End Region
#Region "Acciones"
        Protected MustOverride Function CrearConexion(ByVal Cadena As String) As System.Data.IDbConnection
        Protected MustOverride Function Comando(ByVal ProcedimientoAlmacenado As String) As System.Data.IDbCommand
        Protected MustOverride Function ComandoConsulta(ByVal Consulta As String) As System.Data.IDbCommand
        Protected MustOverride Function CrearDataAdapter(ByVal ProcedimientoAlmacenado As String, ByVal ParamArray Args As Object()) As System.Data.IDataAdapter
        Protected MustOverride Function CrearDataAdapterConsulta(ByVal Consulta As String) As System.Data.IDataAdapter
        Protected MustOverride Function CrearDataAdapter(ByVal ProcedimientoAlmacenado As String, ByVal Parametros() As String, ByVal ParamArray Args As Object()) As System.Data.IDataAdapter
        Protected MustOverride Sub CargarParametros(ByVal Comando As System.Data.IDbCommand, ByVal Args As Object())
        Protected MustOverride Sub CargarParametros(ByVal Comando As System.Data.IDbCommand, ByVal Parametros() As String, ByVal Args As Object())

        ''' <summary> 
        ''' Ejecuta un Procedimiento Almacenado en la base. 
        ''' </summary> 
        Public Function Ejecutar(ByVal ProcedimientoAlmacenado As String) As Integer
            Return Comando(ProcedimientoAlmacenado).ExecuteNonQuery()
        End Function

        Public Function Autenticar() As Boolean
            Try
                If Me.Conexion.State <> System.Data.ConnectionState.Open Then
                    Me.Conexion.Open()
                End If
                Return True
            Catch ex As Exception
                Return False
                Throw New System.Exception(ex.Message)
            End Try
        End Function

        ''' <summary> 
        ''' Ejecuta un Procedimiento Almacenado en la base, utilizando los parámetros. 
        ''' </summary> 
        Public Function Ejecutar(ByVal ProcedimientoAlmacenado As String, ByVal ParamArray Args As Object()) As Integer
            Dim Com As System.Data.IDbCommand = Comando(ProcedimientoAlmacenado)
            CargarParametros(Com, Args)
            Dim Resp As Integer = Com.ExecuteNonQuery()

            For i As Integer = 1 To Com.Parameters.Count - 1
                'For i As Integer = 1 To Com.Parameters.Count - 1 yp,ofifique
                'retrono los valores de la consulta del procedimiento tambien con la cantidad de filas afectadas por defecto
                Dim Par As System.Data.IDbDataParameter = DirectCast(Com.Parameters(i), System.Data.IDbDataParameter)
                If Par.Direction = System.Data.ParameterDirection.InputOutput OrElse Par.Direction = System.Data.ParameterDirection.Output Then
                    Args.SetValue(Par.Value, i - 1)
                    'Args.SetValue(Par.Value, i) yo modifique
                End If
            Next
            Return Resp
        End Function


        Public Function EjecutarOutput(ByVal ProcedimientoAlmacenado As String, ByVal Parametros() As String, ByVal ParamArray Args As Object()) As Object()
            Dim Com As System.Data.IDbCommand = Comando(ProcedimientoAlmacenado)
            Dim Retorno(Parametros.Count - 1) As Object
            CargarParametros(Com, Parametros, Args)
            Dim Resp As Integer = Com.ExecuteNonQuery()

            For i As Integer = 1 To Com.Parameters.Count - 1
                'retrono los valores de la consulta del procedimiento tambien con la cantidad de filas afectadas por defecto
                Dim Par As System.Data.IDbDataParameter = DirectCast(Com.Parameters(i), System.Data.IDbDataParameter)
                If Par.Direction = System.Data.ParameterDirection.InputOutput OrElse Par.Direction = System.Data.ParameterDirection.Output OrElse Par.Direction = System.Data.ParameterDirection.ReturnValue Then
                    'Args.SetValue(Par.Value, i - 1)
                    For c As Integer = 0 To Parametros.Length - 1
                        If Par.ParameterName = Parametros(c) Then
                            Retorno.SetValue(Par.Value, c)
                        End If
                    Next
                End If
            Next

            Return Retorno
        End Function

#End Region

#Region "Transacciones"
        Protected mTransaccion As System.Data.IDbTransaction
        Protected EnTransaccion As Boolean = False
        ''' <summary> 
        ''' Comienza una Transacción en la base en uso. 
        ''' </summary> 
        Public Sub IniciarTransaccion()
            mTransaccion = Me.Conexion.BeginTransaction()
            EnTransaccion = True
        End Sub
        ''' <summary> 
        ''' Confirma la transacción activa. 
        ''' </summary> 
        Public Sub TerminarTransaccion()
            Try
                mTransaccion.Commit()
            Catch Ex As System.Exception
                Throw Ex
            Finally
                mTransaccion = Nothing
                EnTransaccion = False
            End Try
        End Sub
        ''' <summary> 
        ''' Cancela la transacción activa. 
        ''' </summary> 
        Public Sub AbortarTransaccion()
            Try
                mTransaccion.Rollback()
            Catch Ex As System.Exception
                Throw Ex
            Finally
                mTransaccion = Nothing
                EnTransaccion = False
            End Try
        End Sub
#End Region
    End Class
End Namespace

