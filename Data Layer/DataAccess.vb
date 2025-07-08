
Option Explicit On
Option Strict On


Imports System.Configuration.ConfigurationManager
Imports System.Configuration
Imports System.Data.SqlClient

Public Class DataAccess

#Region " Attribute Declarations "

    ' Enumeration of valid return types that can be specidied for a stored procedure 
    Public Enum StoredProcedureReturnType
        DataTable = 1
        RowsAffected = 2
        Scalar = 3
    End Enum

#End Region

#Region " Constructor "

    Private Sub New()
        ' constructor is marked as private so this class 
        ' cannot be instantiated                     
    End Sub

#End Region

#Region " Public Methods "
    ' Returns a Sql Parameter object OVERLOADED 
    ' --------------------------------------------------------------------------
    ' commandText: The text of the query to run.
    Public Shared Function SetSQLParameterProperties( _
        ByVal parameterName As String, _
        ByVal parameterDbType As System.Data.DbType, _
        ByVal parameterValue As Object) As System.Data.SqlClient.SqlParameter

        Dim Parameter As SqlClient.SqlParameter

        Try
            Parameter = New SqlClient.SqlParameter

            Parameter.ParameterName = parameterName
            Parameter.DbType = parameterDbType
            Parameter.Value = parameterValue

            Return Parameter

        Catch ex As Exception
            Throw

        Finally
            Parameter = Nothing

        End Try

    End Function


    ' Creates an open SqlConnection for handling transactions in code
    Public Shared Sub CreateOpenConnection(ByRef sqlConn As SqlClient.SqlConnection)
        Try
            ' Create the open connection
            sqlConn = New System.Data.SqlClient.SqlConnection(ConfigurationManager.ConnectionStrings("Ultipro").ConnectionString) ' <- Connection strings fetched from Config/ConnectionString File
            'AppSettings.Get("UltiproConnectionString"))
            sqlConn.Open()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    'Executes a stored procedure in the specified Database and 
    'returns an object containing the return from the stored procedure 
    '--------------------------------------------------------------------------
    'StoredProcedureName: Name of the stored procedure to execute.
    'ReturnType: Describes the output type expected.
    'MyParameters: Collection of parameter objects for SP.
    Public Shared Function ExecuteStoredProcedure( _
        ByVal storedProcedureName As String, _
        ByVal returnType As StoredProcedureReturnType, _
        ByRef parameters As Collection) As Object
        Try
            ' Create an open connection to the database
            Using mySqlConnection As SqlConnection = New System.Data.SqlClient.SqlConnection(ConfigurationManager.ConnectionStrings("Ultipro").ConnectionString) ' <- Connection strings fetched from Config/ConnectionString File
                'AppSettings.Get("UltiproConnectionString"))
                mySqlConnection.Open()

                ' execute the stored procedure and return the appropriate type 
                Return DataAccess.ExecuteStoredProcedureCommon( _
                    mySqlConnection, _
                    storedProcedureName, _
                    returnType, _
                    parameters, _
                    Nothing)
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Function


    ' Public method containing core stored procedure execution logic that is
    ' indepenMySqlDataAdapternt of whether a transaction has been specified.
    ' --------------------------------------------------------------------------
    ' Connection: An open SQLConnection object.</param>
    ' StoredProcedureName: The name of the stored procedure to run.</param>
    ' MyParameters: A collection of parameters to be passed to the stored procedure.
    ' ReturnType: Describes the type of Data expected to be returned.
    ' --------------------------------------------------------------------------
    ' Returns: An object representing an instance of the ReturnType. 
    Public Shared Function ExecuteStoredProcedureCommon( _
        ByVal connection As System.Data.SqlClient.SqlConnection, _
        ByVal storedProcedureName As String, _
        ByVal returnType As StoredProcedureReturnType, _
        ByRef parameters As Collection, _
        ByRef transaction As System.Data.SqlClient.SqlTransaction) As Object
        Dim iParameterIndex As Integer
        Try
            ' instantiate the command object 
            Using MySqlCommand As SqlCommand = New SqlCommand(storedProcedureName, connection)
                ' set the command parameters 
                MySqlCommand.CommandType = CommandType.StoredProcedure
                MySqlCommand.CommandTimeout = 600
                If parameters IsNot Nothing Then
                    ' Add command parameters 
                    For iParameterIndex = 1 To (parameters.Count)
                        If parameters IsNot Nothing Then
                            MySqlCommand.Parameters.Add(parameters.Item(iParameterIndex))
                        End If
                    Next
                End If

                ' if a transaction object exists, associate it with the command object 
                If Not transaction Is Nothing Then
                    MySqlCommand.Transaction = transaction
                End If

                ' execute the stored procedure and return the appropriate type 
                Select Case returnType
                    Case StoredProcedureReturnType.DataTable
                        ' set Data aMySqlDataAdapterpter parameters 
                        Using MySqlDataAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(MySqlCommand)
                            ' fill the Datatable with the query results 
                            Using MyDataTable As DataTable = New DataTable
                                MySqlDataAdapter.Fill(MyDataTable)
                                Return MyDataTable
                            End Using
                        End Using
                    Case StoredProcedureReturnType.RowsAffected
                        Return MySqlCommand.ExecuteNonQuery()
                    Case StoredProcedureReturnType.Scalar
                        Return MySqlCommand.ExecuteScalar()
                    Case Else
                        Return Nothing

                End Select
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region


End Class

