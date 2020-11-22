#Const Support_Access = True
#Const Support_MSSQL = True
#Const Support_MySQL = False
#Const Support_Oracle = False
#Const Support_SQLite = False
#Const Support_TXT = False
#Const Support_MSSQL_Compact = False

Imports System.Threading
Imports System.Threading.Thread
Imports System.Linq

Public Module ADODB
    'replacement for built in UBound guaranteed to never throw an error
    Private Function UBound(ByVal obj As System.Array, Optional Rank As Integer = 1) As Integer
        Try
            If obj Is Nothing Then Return -1
            Return obj.GetUpperBound(Rank - 1)
        Catch ex As Exception
            Return -1
        End Try
    End Function

    'replacement for built in UBound guaranteed to never throw an error
    Private Function UBound(ByVal obj As Collections.ICollection) As Integer
        Try
            If obj Is Nothing Then Return -1
            Return obj.Count - 1
        Catch
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Turns a string array into a SQL nest - such as {"Tom", "Fred", "Bob's"} = "'Tom','Fred','Bob''s'", taking injection into consideration
    ''' </summary>
    ''' <param name="source">String array</param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()> Public Function SQLNest(ByVal source() As String) As String
        Dim count = UBound(source)
        If count > -1 Then
            Dim ret(count) As String
            For i As Integer = 0 To count
                If source(i) IsNot Nothing Then
                    ret(i) = "'" & source(i).Replace("'", "''") & "'" 'TODO: does not take MySQL backslashes into consideration
                End If
            Next
            Return String.Join(",", ret)
        Else
            Return "''" 'prevent SQL error, but will never match if columns don't allow empty strings
        End If
    End Function

    ''' <summary>
    ''' Turns a string array into a SQL nest - such as {"Tom", "Fred", "Bob's"} = "'Tom','Fred','Bob''s'", taking injection into consideration
    ''' </summary>
    ''' <param name="source">String list</param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()> Public Function SQLNest(ByVal source As List(Of String)) As String
        Dim count = UBound(source)
        If count > -1 Then
            Dim ret(count) As String
            For i As Integer = 0 To count
                If source(i) IsNot Nothing Then
                    ret(i) = "'" & source(i).Replace("'", "''") & "'" 'TODO: does not take MySQL backslashes into consideration
                End If
            Next
            Return String.Join(",", ret)
        Else
            Return "''" 'prevent SQL error, but will never match if columns don't allow empty strings
        End If
    End Function

    'Turns an array into a SQL nest - such as {0,1,2,3} = 0,1,2,3
    <System.Runtime.CompilerServices.Extension()> Public Function SQLNest(Of T)(ByVal source As T()) As String
        If UBound(source) = -1 Then Return Nothing
        Return String.Join(",", source)
    End Function

    'Turns an enum into a SQL nest - such as {0,1,2,3} = 0,1,2,3
    <System.Runtime.CompilerServices.Extension()> Public Function SQLNest(Of T)(ByVal source As IEnumerable(Of T)) As String
        If source Is Nothing OrElse Not source.Any(Function(i) True) Then Return Nothing
        Return String.Join(",", source)
    End Function

    ''' <summary>
    ''' Wrap this class in a using( ) { } statement around database queries where it a cancellation token is needed to stop a SQL command by calling IDbCommand.Cancel().
    ''' </summary>
    Private Class RunSQLCancellable
        Implements IDisposable

        Private _cancellationTokenRegistration As CancellationTokenRegistration

        ''' <summary>
        ''' </summary>
        ''' <param name="cancelDbCommand">IDbCommand to call .Cancel() on if the token is cancelled</param>
        ''' <param name="token">CancellationToken to register</param>
        Public Sub New(cancelDbCommand As Data.IDbCommand, token As CancellationToken)
            If token <> Nothing Then
                If token.IsCancellationRequested Then Throw New Exception("Cancellation Token Has Requested Abort")
                Me._cancellationTokenRegistration = token.Register(Sub() cancelDbCommand.Cancel())
            End If
        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    If Me._cancellationTokenRegistration <> Nothing Then
                        Me._cancellationTokenRegistration.Dispose()
                        Me._cancellationTokenRegistration = Nothing
                    End If
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            ' TODO: uncomment the following line if Finalize() is overridden above.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class

    Private ACCESS_DATABASE_ENGINE_X64 As String = "Microsoft.ACE.OLEDB.12.0"

    Public Enum LockTypeEnum
        adLockUnspecified = -1
        adLockBatchOptimistic = 4
        adLockOptimistic = 3
        adLockPessimistic = 2
        adLockReadOnly = 1
    End Enum

    Public Enum ObjectStateEnum
        adStateClosed = 0
        adStateOpen = 1
        adStateConnecting = 2
        adStateExecuting = 4
        adStateFetching = 8
    End Enum

    Public Enum EditModeEnum
        adEditNone = 0
        adEditInProgress = 1
        adEditAdd = 2
        adEditDelete = 4
    End Enum

    Public Enum CursorLocationEnum
        adUseClient = 3
        adUseServer = 2
    End Enum

    Public Enum PositionEnum
        adPosBOF = -2
        adPosEOF = -3
        adPosUnknown = -1
    End Enum

    Public Enum CursorTypeEnum
        ''' <summary>
        ''' Unspecified will force a client side cursor
        ''' </summary>
        adOpenUnspecified = -1
        adOpenDynamic = 2
        adOpenForwardOnly = 0
        adOpenKeyset = 1
        adOpenStatic = 3
    End Enum

    Public Enum DataTypeEnum
        adBigInt = 20
        adBinary = 128
        adBoolean = 11
        adBSTR = 8
        adChapter = 136
        adChar = 129
        adCurrency = 6
        adDate = 7
        adDBDate = 133
        adDBFileTime = 137
        adDBTime = 134
        adDBTimeStamp = 135
        adDBTimeStampOffset = 146
        adDecimal = 14
        adDouble = 5
        adEmpty = 0
        adError = 10
        adFileTime = 64
        adGUID = 72
        adIDispatch = 9
        adInteger = 3
        adIUnknown = 13
        adLongVarBinary = 205
        adLongVarChar = 201
        adLongVarWChar = 203
        adNumeric = 131
        adPropVariant = 138
        adSingle = 4
        adSmallInt = 2
        adTinyInt = 16
        adUnsignedBigInt = 21
        adUnsignedInt = 19
        adUnsignedSmallInt = 18
        adUnsignedTinyInt = 17
        adUserDefined = 132
        adVarBinary = 204
        adVarChar = 200
        adVariant = 12
        adVarNumeric = 139
        adVarWChar = 202
        adWChar = 130
    End Enum

    Public Enum CommandTypeEnum
        adCmdUnspecified = -1
        adCmdText = 1
        adCmdTable = 2
        adCmdStoredProc = 4
        adCmdUnknown = 8
        adCmdFile = 256
        adCmdTableDirect = 512
    End Enum

    Public Enum ParameterDirectionEnum
        adParamUnknown = 0
        adParamInput = 1
        adParamOutput = 2
        adParamInputOutput = 3
        adParamReturnValue = 4
    End Enum

    Public Enum SearchDirectionEnum
        adSearchForward = 1
        adSearchBackward = -1
    End Enum

    Public Enum ConnectOptionEnum
        adConnectUnspecific = -1
        adAsyncConnect = 16
    End Enum

    Public Enum ExecuteOptionEnum
        adOptionUnspecified = -1
        'adAsyncExecute = 16 'no support for async yet
        'adAsyncFetch = 32
        'adAsyncFetchNonBlocking = 64
        adExecuteNoRecords = 128
        adExecuteGetReturnValue = 256
        'adExecuteStream = 1024 'no support for streaming yet
        'adExecuteRecord = 2048 'use Open( ) if you want a recordset
    End Enum

    Public Enum eErrorType
        ERROR_OUT_OF_BOUNDS = 3021
        ERROR_CONTEXT = 3219
        ERROR_NOUPDATES = 3251
        ERROR_NOT_FOUND = 3265
        ERROR_CLOSED = 3704
        ERROR_OPEN = 3705
        ERROR_CONNECTION_CLOSED = 3709
        ERROR_NO_TRANS = &H8004D00E
    End Enum

    Public Enum enumProvider
        provider_AUTO = -1
        provider_MSACCESS = 0
        provider_SQLSERVER = 1
        provider_ORACLE = 2
        provider_MYSQL = 3
        provider_SQLITE = 4
        provider_SQLSERVER_COMPACT = 5
        provider_TXT = 99
    End Enum

    Public Enum SchemaEnum
        adSchemaProviderSpecific = -1
        adSchemaAsserts = 0
        adSchemaCatalogs = 1
        adSchemaCharacterSets = 2
        adSchemaCollations = 3
        adSchemaColumns = 4
        adSchemaCheckConstraints = 5
        adSchemaConstraintColumnUsage = 6
        adSchemaConstraintTableUsage = 7
        adSchemaKeyColumnUsage = 8
        adSchemaReferentialContraints = 9 'yes, this spelling mistake is in the original COM object
        adSchemaReferentialConstraints = 9
        adSchemaTableConstraints = 10
        adSchemaColumnsDomainUsage = 11
        adSchemaIndexes = 12
        adSchemaColumnPrivileges = 13
        adSchemaTablePrivileges = 14
        adSchemaUsagePrivileges = 15
        adSchemaProcedures = 16
        adSchemaSchemata = 17
        adSchemaSQLLanguages = 18
        adSchemaStatistics = 19
        adSchemaTables = 20
        adSchemaTranslations = 21
        adSchemaProviderTypes = 22
        adSchemaViews = 23
        adSchemaViewColumnUsage = 24
        adSchemaViewTableUsage = 25
        adSchemaProcedureParameters = 26
        adSchemaForeignKeys = 27
        adSchemaPrimaryKeys = 28
        adSchemaProcedureColumns = 29
        adSchemaDBInfoKeywords = 30
        adSchemaDBInfoLiterals = 31
        adSchemaCubes = 32
        adSchemaDimensions = 33
        adSchemaHierarchies = 34
        adSchemaLevels = 35
        adSchemaMeasures = 36
        adSchemaProperties = 37
        adSchemaMembers = 38
        adSchemaTrustees = 39
        adSchemaFunctions = 40
        adSchemaActions = 41
        adSchemaCommands = 42
        adSchemaSets = 43
    End Enum

    Public Enum StringFormatEnum
        adClipString = 2
    End Enum

    Public Class Connection
        Implements IDisposable

        Private Structure pair
            Public Name As String
            Public Value As String
        End Structure

        Private Structure strucConnectionString
            Public PROVIDER As pair
            Public DRIVER As pair
            Public NETWORK As pair
            Public SERVER As pair
            Public DATABASE As pair
            Public UID As pair
            Public PWD As pair
            Public TIMEOUT As pair
            Public DSN As pair
            Public DATASOURCE As pair
            Public EXTENSIONS As pair
            Public DBQ As pair
            Public [OPTION] As pair
            Public EXTENDEDPROPERTIES As pair
            Public APPLICATIONNAME As pair
            Public MARS As pair
            Public ProviderType As enumProvider
            Public CONNECTIONSTRING As String
        End Structure

        Public connection As Data.Common.DbConnection
        Private m_ConnectionString As String
        Private m_connectiondata As strucConnectionString
        Protected Friend connection_transaction As Data.IDbTransaction
        Private m_CommandTimeout As Integer
        Protected Friend m_Errors As Exception
        Private m_ConnectionTimeout As Integer
        Private m_CursorLocation As CursorLocationEnum
        Private m_CursorType As CursorTypeEnum
        Private m_State As ObjectStateEnum
        Private m_username As String
        Private m_password As String
        Private m_sqldata As enumProvider
        Private m_pooling As Boolean
        Private m_application As String
        Private trans_counter As Integer

        ''' <summary>
        ''' If this is set to true, calls to DB.Executes("INSERT/DELETE..."),
        ''' along with Recordsets opened with this Connection,
        ''' will not be able to Update or Delete records
        ''' </summary>
        Public DisableUpdates As Boolean

        ''' <summary>
        ''' Set this to a Cancellation Token if you need to cancel a running query when it is triggered by the source
        ''' </summary>
        Public CancellationToken As CancellationToken

        Protected Friend ChildObjects As New List(Of Recordset) 'list of Recordsets's that are referencing us

        Public Sub BeginTrans()
            If Interlocked.Increment(trans_counter) = 1 Then
                connection_transaction = connection.BeginTransaction
            End If
        End Sub

        Public Sub Close()
            If m_State = ObjectStateEnum.adStateClosed Then DoError(eErrorType.ERROR_CLOSED)

            'For i As Integer = 0 To ChildObjects.Count - 1
            'Try
            'ChildObjects(i).Dispose()
            'Catch
            'End Try
            'Next
            ChildObjects.Clear()

            Connection_Cleanup()

            'because this object can be reused, reinitialize approriate variables here
            m_ConnectionString = ""
            m_CommandTimeout = 0
            m_ConnectionTimeout = 0
            m_CursorLocation = CursorLocationEnum.adUseServer
            m_CursorType = CursorTypeEnum.adOpenUnspecified
            m_State = ObjectStateEnum.adStateClosed
            m_username = ""
            m_password = ""
            m_sqldata = -1
            m_pooling = False
            m_application = ""
            m_State = ObjectStateEnum.adStateClosed
        End Sub

        Public ReadOnly Property Errors() As Errors
            Get
                Return New Errors() With {.conn = Me}
            End Get
        End Property

        Public Property ConnectionPooling() As Boolean
            Get
                Return m_pooling
            End Get
            Set(ByVal value As Boolean)
                m_pooling = False
            End Set
        End Property

        Public Property Application() As String
            Get
                Return m_application
            End Get
            Set(ByVal value As String)
                m_application = value
            End Set
        End Property

        Public Property CommandTimeout() As Integer
            Get
                Return m_CommandTimeout
            End Get
            Set(ByVal Value As Integer)
                m_CommandTimeout = Value
            End Set
        End Property

        Public Property ConnectionTimeout() As Integer
            Get
                Return m_ConnectionTimeout
            End Get
            Set(ByVal Value As Integer)
                m_ConnectionTimeout = Value
            End Set
        End Property

        Public Sub CommitTrans()
            If Interlocked.Decrement(trans_counter) <= 0 Then
                If trans_counter < 0 Then
                    Interlocked.Increment(trans_counter)
                    DoError(eErrorType.ERROR_NO_TRANS)
                Else
                    connection_transaction.Commit()
                    For Each RS As Recordset In ChildObjects
                        RS.data_table.AcceptChanges()
                    Next
                    connection_transaction = Nothing
                    trans_counter = 0
                End If
            End If
        End Sub

        Public Property ConnectionString() As String
            Get
                Return m_ConnectionString
            End Get
            Set(ByVal Value As String)
                m_ConnectionString = Value
            End Set
        End Property

        Public Property CursorLocation() As CursorLocationEnum
            Get
                Return m_CursorLocation
            End Get
            Set(ByVal Value As CursorLocationEnum)
                m_CursorLocation = Value
            End Set
        End Property

        Private ReadOnly Property ProviderType() As enumProvider
            Get
                Return m_connectiondata.ProviderType
            End Get
        End Property

        Public ReadOnly Property Provider() As String
            Get
                Return m_connectiondata.PROVIDER.Value
            End Get
        End Property

        Public ReadOnly Property DefaultDatabase() As String
            Get
                Return m_connectiondata.DATABASE.Value
            End Get
        End Property

        ''' <summary>
        ''' Acquire an application lock that has the scope of the current database.
        ''' This lock will automatically release when ReleaseAppLock is called, or the connection is terminated.
        ''' !BE CAREFUL, THIS CAN BE RECURSED AND YOU HAVE TO CALL RELEASE THE SAME NUMBER OF TIMES THAT YOU CALL ACQUIRE!
        ''' </summary>
        ''' <param name="name">Name of lock</param>
        ''' <param name="timeout">seconds to wait for lock (0=no wait)</param>
        ''' <returns>true/false if the lock was obtained</returns>
        Public Function AcquireAppLock(Optional name As String = "applock", Optional timeout As Integer = 120000) As Boolean
            Try
                If Me.Execute("EXECUTE sp_getapplock @Resource = '" & name.Replace("'", "''") & "', @LockMode = 'Exclusive', @LockTimeout = " & timeout & ", @LockOwner='Session'", Options:=ExecuteOptionEnum.adExecuteGetReturnValue) >= 0 Then
                    Return True
                End If
            Catch
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Release previously acquired application lock
        ''' </summary>
        ''' <param name="name"></param>
        Public Function ReleaseAppLock(Optional name As String = "applock") As Boolean
            Try
                If Me.Execute("EXECUTE sp_releaseapplock @Resource = '" & name.Replace("'", "''") & "', @LockOwner='Session'", Options:=ExecuteOptionEnum.adExecuteGetReturnValue) >= 0 Then
                    Return True
                End If
            Catch
            End Try

            Return False
        End Function

        Public Function Execute(ByVal CommandText As String, Optional ByRef RecordsAffected As Integer = 0, Optional ByVal Options As ExecuteOptionEnum = ExecuteOptionEnum.adExecuteNoRecords) As Integer
            If Me.DisableUpdates Then Return 0

            Dim getReturnValue As Boolean = False
            If ((Options And ExecuteOptionEnum.adExecuteGetReturnValue) = ExecuteOptionEnum.adExecuteGetReturnValue AndAlso TypeOf connection Is Data.SqlClient.SqlConnection) Then getReturnValue = True

            Using adocommand = Get_ProviderSpecific_Command()
                adocommand.CommandTimeout = m_CommandTimeout
                adocommand.CommandText = CommandText
                adocommand.Connection = connection
                If Not connection_transaction Is Nothing Then adocommand.Transaction = connection_transaction

                'pull return value for SQL
                Dim returnValue As Data.IDbDataParameter = Nothing
                If getReturnValue Then
                    returnValue = adocommand.CreateParameter()
                    returnValue.Direction = Data.ParameterDirection.ReturnValue
                    adocommand.Parameters.Add(returnValue)
                End If

                Using New RunSQLCancellable(adocommand, Me.CancellationToken)
                    RecordsAffected = adocommand.ExecuteNonQuery()
                End Using

                Return If(Not returnValue Is Nothing, Val(returnValue.Value), 0)
            End Using
        End Function

        Public ReadOnly Property Properties() As Properties
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Public Sub Open(Optional ByVal ConnectionString As String = "", Optional ByVal UserID As String = "", Optional ByVal Password As String = "", Optional ByVal Options As ConnectOptionEnum = ConnectOptionEnum.adConnectUnspecific, Optional ByVal providertype As enumProvider = enumProvider.provider_AUTO)
            Try
                If Len(UserID) > 0 Then m_username = UserID
                If Len(Password) > 0 Then m_password = Password

                If Len(ConnectionString) > 0 Then m_ConnectionString = ConnectionString
                Prepare_Connection_String(providertype)

                Select Case m_connectiondata.ProviderType
#If Support_MSSQL Then
                    Case enumProvider.provider_SQLSERVER
                        connection = New Data.SqlClient.SqlConnection
#End If
#If Support_MSSQL_Compact Then
                    Case enumProvider.provider_SQLSERVER_COMPACT
                        connection = New Data.SqlServerCe.SqlCeConnection
#End If
#If Support_MySQL Then
                    Case enumProvider.provider_MYSQL
                        connection = New MySql.Data.MySqlClient.MySqlConnection
#End If
#If Support_Oracle Then
                    Case enumProvider.provider_ORACLE
                        connection = New Data.OracleClient.OracleConnection
#End If
#If Support_SQLite Then
                    Case enumProvider.provider_SQLITE
                        connection = New Data.SQLite.SQLiteConnection
#End If
                    Case Else
                        connection = New Data.OleDb.OleDbConnection
                End Select

                connection.ConnectionString = m_connectiondata.CONNECTIONSTRING

                If Options = ConnectOptionEnum.adAsyncConnect Then
                    Dim t = New Thread(AddressOf Open_Async)
                    t.Start()
                Else
                    m_State = ObjectStateEnum.adStateConnecting
                    connection.Open()
                    m_State = ObjectStateEnum.adStateOpen
                End If
            Catch ex As Exception
                '#If Support_Access Then
                '                If Environment.Is64BitProcess AndAlso InStr(1, ex.Message, "The 'Microsoft.ACE.OLEDB.16.0' provider is not registered on the local machine.") > 0 Then
                '                    'switch to ACE 12 if 16 is not installed and try again, note that 12 will crash occassionally
                '                    ACCESS_DATABASE_ENGINE_X64 = "Microsoft.ACE.OLEDB.12.0"
                '                    Open(ConnectionString, UserID, Password, Options, providertype)
                '                    Return
                '                End If
                '#End If
                m_Errors = ex
                Throw
            End Try
        End Sub

        'Used for the adAsyncConnect option
        Private Sub Open_Async()
            Try
                m_State = ObjectStateEnum.adStateConnecting
                connection.Open()
                m_State = ObjectStateEnum.adStateOpen
            Catch ex As Exception
                m_Errors = ex
                m_State = ObjectStateEnum.adStateClosed
            End Try
        End Sub

        Public Function OpenSchema(ByVal Schema As SchemaEnum, Optional ByVal Restrictions() As Object = Nothing) As Recordset
            Dim schema_guid As Guid
            Select Case Schema
                Case SchemaEnum.adSchemaProviderSpecific
                    Throw New NotImplementedException
                    'schema_guid = Data.OleDb.OleDbSchemaGuid.
                Case SchemaEnum.adSchemaAsserts
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Assertions
                Case SchemaEnum.adSchemaCatalogs
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Catalogs
                Case SchemaEnum.adSchemaCharacterSets
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Character_Sets
                Case SchemaEnum.adSchemaCollations
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Collations
                Case SchemaEnum.adSchemaColumns
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Columns
                Case SchemaEnum.adSchemaCheckConstraints
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Check_Constraints
                Case SchemaEnum.adSchemaConstraintColumnUsage
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Constraint_Column_Usage
                Case SchemaEnum.adSchemaConstraintTableUsage
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Constraint_Table_Usage
                Case SchemaEnum.adSchemaKeyColumnUsage
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Key_Column_Usage
                Case SchemaEnum.adSchemaReferentialContraints, SchemaEnum.adSchemaReferentialConstraints
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Referential_Constraints
                Case SchemaEnum.adSchemaTableConstraints
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Table_Constraints
                Case SchemaEnum.adSchemaColumnsDomainUsage
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Column_Domain_Usage
                Case SchemaEnum.adSchemaIndexes
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Indexes
                Case SchemaEnum.adSchemaColumnPrivileges
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Column_Privileges
                Case SchemaEnum.adSchemaTablePrivileges
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Table_Privileges
                Case SchemaEnum.adSchemaUsagePrivileges
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Usage_Privileges
                Case SchemaEnum.adSchemaProcedures
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Procedures
                Case SchemaEnum.adSchemaSchemata
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Schemata
                Case SchemaEnum.adSchemaSQLLanguages
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Sql_Languages
                Case SchemaEnum.adSchemaStatistics
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Statistics
                Case SchemaEnum.adSchemaTables
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Tables
                Case SchemaEnum.adSchemaTranslations
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Translations
                Case SchemaEnum.adSchemaProviderTypes
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Provider_Types
                Case SchemaEnum.adSchemaViews
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Views
                Case SchemaEnum.adSchemaViewColumnUsage
                    schema_guid = Data.OleDb.OleDbSchemaGuid.View_Column_Usage
                Case SchemaEnum.adSchemaViewTableUsage
                    schema_guid = Data.OleDb.OleDbSchemaGuid.View_Table_Usage
                Case SchemaEnum.adSchemaProcedureParameters
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Procedure_Parameters
                Case SchemaEnum.adSchemaForeignKeys
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Foreign_Keys
                Case SchemaEnum.adSchemaPrimaryKeys
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Primary_Keys
                Case SchemaEnum.adSchemaProcedureColumns
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Procedure_Columns
                Case SchemaEnum.adSchemaDBInfoKeywords
                    schema_guid = Data.OleDb.OleDbSchemaGuid.DbInfoKeywords
                Case SchemaEnum.adSchemaDBInfoLiterals
                    schema_guid = Data.OleDb.OleDbSchemaGuid.DbInfoLiterals
                Case SchemaEnum.adSchemaCubes
                    Throw New NotImplementedException
                    'schema_guid = Data.OleDb.OleDbSchemaGuid.
                Case SchemaEnum.adSchemaDimensions
                    Throw New NotImplementedException
                    'schema_guid = Data.OleDb.OleDbSchemaGuid.
                Case SchemaEnum.adSchemaHierarchies
                    Throw New NotImplementedException
                    'schema_guid = Data.OleDb.OleDbSchemaGuid.
                Case SchemaEnum.adSchemaLevels
                    Throw New NotImplementedException
                    'schema_guid = Data.OleDb.OleDbSchemaGuid.
                Case SchemaEnum.adSchemaMeasures
                    Throw New NotImplementedException
                    'schema_guid = Data.OleDb.OleDbSchemaGuid.
                Case SchemaEnum.adSchemaProperties
                    Throw New NotImplementedException
                    'schema_guid = Data.OleDb.OleDbSchemaGuid.
                Case SchemaEnum.adSchemaMembers
                    Throw New NotImplementedException
                    'schema_guid = Data.OleDb.OleDbSchemaGuid.
                Case SchemaEnum.adSchemaTrustees
                    schema_guid = Data.OleDb.OleDbSchemaGuid.Trustee
                Case SchemaEnum.adSchemaFunctions
                    Throw New NotImplementedException
                    'schema_guid = Data.OleDb.OleDbSchemaGuid.
                Case SchemaEnum.adSchemaActions
                    Throw New NotImplementedException
                    'schema_guid = Data.OleDb.OleDbSchemaGuid.
                Case SchemaEnum.adSchemaCommands
                    schema_guid = Data.OleDb.OleDbSchemaGuid.SchemaGuids
            End Select

            Dim tmp_conn As New Data.OleDb.OleDbConnection(m_ConnectionString)
            tmp_conn.Open()
            'tmp_conn.GetOleDbSchemaTable(schema_guid, Restrictions)
            Dim RS As New Recordset()
            RS.InitOpen(tmp_conn.GetOleDbSchemaTable(schema_guid, Restrictions), Me)
            tmp_conn.Close()
            tmp_conn.Dispose()
            Return RS
        End Function

        Public Sub RollbackTrans()
            If trans_counter = 0 Then
                DoError(eErrorType.ERROR_NO_TRANS)
            Else
                connection_transaction.Rollback()
                For Each RS As Recordset In ChildObjects
                    RS.CancelUpdate()
                    RS.data_table.RejectChanges()
                Next
                connection_transaction = Nothing
                trans_counter = 0
            End If
        End Sub

        Public ReadOnly Property State() As ObjectStateEnum
            Get
                Return m_State
            End Get
        End Property

        Public Property UserName() As String
            Get
                Return m_username
            End Get
            Set(ByVal value As String)
                m_username = value
            End Set
        End Property

        Public Property Password() As String
            Get
                Return m_password
            End Get
            Set(ByVal value As String)
                m_password = value
            End Set
        End Property

        Public Function Version() As String
            If False Then
                Return Nothing
#If Support_MSSQL Then
            ElseIf TypeOf connection Is Data.SqlClient.SqlConnection Then
                Return CType(connection, Data.SqlClient.SqlConnection).ServerVersion
#End If
#If Support_MSSQL_Compact Then
            ElseIf TypeOf connection Is Data.SqlServerCe.SqlCeConnection Then
                Return CType(connection, Data.SqlServerCe.SqlCeConnection).ServerVersion
#End If
#If Support_MySQL Then
            ElseIf TypeOf connection Is MySql.Data.MySqlClient.MySqlConnection Then
                Return CType(connection, MySql.Data.MySqlClient.MySqlConnection).ServerVersion
#End If
#If Support_Oracle Then
            ElseIf TypeOf connection Is Data.OracleClient.OracleConnection Then
                Return CType(connection, Data.OracleClient.OracleConnection).ServerVersion
#End If
#If Support_SQLite Then
            ElseIf TypeOf connection Is Data.Sqlite.SqliteConnection Then
                Return CType(connection, Data.Sqlite.SqliteConnection).ServerVersion
#End If
            Else
                Return CType(connection, Data.OleDb.OleDbConnection).ServerVersion
            End If
        End Function

        'Protected Friend ReadOnly Property oledb_connection() As Data.IDbConnection
        'Get
        'Return connection
        'End Get
        'End Property

        Public Sub New()
            'defaults
            m_State = ObjectStateEnum.adStateClosed
            m_CursorLocation = CursorLocationEnum.adUseServer
            'm_pooling = True
        End Sub

        Private Sub Connection_Cleanup()
            If Not connection_transaction Is Nothing Then connection_transaction.Dispose()
            If Not connection Is Nothing Then
                Try
                    connection.Close()
                Catch
                End Try
                connection.Dispose()
            End If

            'clean up objects
            connection = Nothing
            connection_transaction = Nothing
            m_connectiondata = Nothing
        End Sub

        Private Sub Prepare_Connection_String(ByVal providertype As enumProvider)
            Dim connectarr() As String
            Dim i As Integer = 0
            Dim pos As Integer = 0
            Dim itemname As String = ""
            Dim itemvalue As String = ""
            Dim buildconnection(10) As String

            connectarr = Split(m_ConnectionString, ";")
            For i = 0 To UBound(connectarr)
                If Len(connectarr(i)) = 0 Then Continue For
                If InStr(1, connectarr(i), """") > 0 Then 'combine connection properties that are encased in quotes
                    Dim j = i + 1
                    Do While j < UBound(connectarr)
                        Dim str = connectarr(j)
                        connectarr(j) = ""
                        connectarr(i) &= ";" & str
                        j += 1
                        If InStr(1, str, """") > 0 Then Exit Do
                    Loop
                End If

                pos = InStr(1, connectarr(i), "=")
                If pos > 0 Then
                    itemname = UCase(Left(connectarr(i), pos - 1))
                    itemvalue = Mid(connectarr(i), pos + 1)

                    Select Case itemname
                        Case "PROVIDER"
                            m_connectiondata.PROVIDER.Name = itemname
                            m_connectiondata.PROVIDER.Value = UCase(itemvalue)
                        Case "DRIVER"
                            m_connectiondata.DRIVER.Name = itemname
                            m_connectiondata.DRIVER.Value = itemvalue
                        Case "NETWORK"
                            m_connectiondata.NETWORK.Name = itemname
                            m_connectiondata.NETWORK.Value = itemvalue
                        Case "SERVER"
                            m_connectiondata.SERVER.Name = itemname
                            m_connectiondata.SERVER.Value = itemvalue
                        Case "DATABASE"
                            m_connectiondata.DATABASE.Name = itemname
                            m_connectiondata.DATABASE.Value = itemvalue
                        Case "UID", "TRUSTED_CONNECTION"
                            m_connectiondata.UID.Name = itemname
                            m_connectiondata.UID.Value = itemvalue
                        Case "PWD"
                            m_connectiondata.PWD.Name = itemname
                            m_connectiondata.PWD.Value = itemvalue
                        Case "CONNECT TIMEOUT"
                            m_connectiondata.TIMEOUT.Name = itemname
                            m_connectiondata.TIMEOUT.Value = itemvalue
                        Case "DSN"
                            m_connectiondata.DSN.Name = itemname
                            m_connectiondata.DSN.Value = itemvalue
                        Case "DBQ"
                            m_connectiondata.DBQ.Name = itemname
                            m_connectiondata.DBQ.Value = itemvalue
                        Case "DATA SOURCE"
                            m_connectiondata.DATASOURCE.Name = itemname
                            m_connectiondata.DATASOURCE.Value = itemvalue
                        Case "EXTENSIONS"
                            m_connectiondata.EXTENSIONS.Name = itemname
                            m_connectiondata.EXTENSIONS.Value = itemvalue
                        Case "OPTION"
                            m_connectiondata.OPTION.Name = itemname
                            m_connectiondata.OPTION.Value = itemvalue
                        Case "EXTENDED PROPERTIES"
                            m_connectiondata.EXTENDEDPROPERTIES.Name = itemname
                            m_connectiondata.EXTENDEDPROPERTIES.Value = itemvalue
                        Case "APPLICATION NAME"
                            m_connectiondata.APPLICATIONNAME.Name = itemname
                            m_connectiondata.APPLICATIONNAME.Value = itemvalue
                        Case "MULTIPLEACTIVERESULTSETS"
                            m_connectiondata.MARS.Name = itemname
                            m_connectiondata.MARS.Value = itemvalue
                    End Select
                End If
            Next


#If Support_MSSQL Then
            If InStr(1, m_connectiondata.PROVIDER.Value, "sqloledb", Microsoft.VisualBasic.CompareMethod.Text) > 0 OrElse InStr(1, m_connectiondata.DRIVER.Value, "{SQL SERVER", Microsoft.VisualBasic.CompareMethod.Text) > 0 Then
                m_connectiondata.PROVIDER.Name = "PROVIDER"
                m_connectiondata.PROVIDER.Value = "sqloledb"
                m_connectiondata.ProviderType = enumProvider.provider_SQLSERVER
                GoTo done_provider_check
            End If
#End If


#If Support_Access Or Support_MSSQL_Compact Then
            If Len(m_connectiondata.DRIVER.Value) = 0 AndAlso Len(m_connectiondata.PROVIDER.Value) = 0 Then
                m_connectiondata.DATASOURCE.Name = "DATA SOURCE"
                If Len(m_connectiondata.DBQ.Value) > 0 Then m_connectiondata.DATASOURCE.Value = m_connectiondata.DBQ.Value
                If Len(m_connectiondata.DATASOURCE.Value) = 0 Then m_connectiondata.DATASOURCE.Value = m_ConnectionString
#If Support_MSSQL_Compact Then
                If InStrRev(m_connectiondata.DATASOURCE.Value, ".sdf", vbTextCompare) > 0 Then
                    m_connectiondata.ProviderType = enumProvider.provider_SQLSERVER_COMPACT
#Else
                If False Then
#End If
#If Support_Access Then
                Else
                    m_connectiondata.PROVIDER.Name = "PROVIDER"
                    If Environment.Is64BitProcess Then
                        m_connectiondata.PROVIDER.Value = ACCESS_DATABASE_ENGINE_X64
                    Else
                        m_connectiondata.PROVIDER.Value = "Microsoft.Jet.OLEDB.4.0"
                    End If
                    m_connectiondata.ProviderType = enumProvider.provider_MSACCESS
                    GoTo done_provider_check
                End If
            Else
                If InStr(1, m_connectiondata.DRIVER.Value, "Access Driver", Microsoft.VisualBasic.CompareMethod.Text) > 0 _
                    OrElse InStr(1, m_connectiondata.PROVIDER.Value, "Jet.OLEDB", Microsoft.VisualBasic.CompareMethod.Text) > 0 _
                    OrElse InStr(1, m_connectiondata.PROVIDER.Value, "ACE.OLEDB", Microsoft.VisualBasic.CompareMethod.Text) > 0 Then
                    m_connectiondata.PROVIDER.Name = "PROVIDER"
                    If Environment.Is64BitProcess Then
                        m_connectiondata.PROVIDER.Value = ACCESS_DATABASE_ENGINE_X64
                    Else
                        m_connectiondata.PROVIDER.Value = "Microsoft.Jet.OLEDB.4.0"
                    End If
                    m_connectiondata.DATASOURCE.Name = "DATA SOURCE"
                    If Len(m_connectiondata.DBQ.Value) > 0 Then m_connectiondata.DATASOURCE.Value = m_connectiondata.DBQ.Value
                    m_connectiondata.ProviderType = enumProvider.provider_MSACCESS
                    GoTo done_provider_check
#End If
                End If
            End If
#End If


#If Support_MySQL Then
            If InStr(1, m_connectiondata.PROVIDER.Value, "OleMySql", Microsoft.VisualBasic.CompareMethod.Text) > 0 OrElse InStr(1, m_connectiondata.DRIVER.Value, "MySQL", Microsoft.VisualBasic.CompareMethod.Text) > 0 Then
                m_connectiondata.ProviderType = enumProvider.provider_MYSQL
            End If
#End If


#If Support_Oracle Then
            If InStr(1, m_connectiondata.DATASOURCE.Value, "Oracle", Microsoft.VisualBasic.CompareMethod.Text) > 0 orelse InStr(1, m_connectiondata.DATASOURCE.Value, "TORLC", Microsoft.VisualBasic.CompareMethod.Text) > 0 Then
                m_connectiondata.ProviderType = enumProvider.provider_ORACLE
            End If
#End If

#If Support_SQLite Then
            If InStr(1, m_connectiondata.PROVIDER.Value, "SQLite", Microsoft.VisualBasic.CompareMethod.Text) > 0 Then
                m_connectiondata.ProviderType = enumProvider.provider_SQLITE
            End If
#End If

#If Support_TXT Then
            If InStr(1, m_connectiondata.DRIVER.Value, "{Microsoft Text Driver", Microsoft.VisualBasic.CompareMethod.Text) > 0 Then
                m_connectiondata.ProviderType = enumProvider.provider_TXT
            End If
#End If

done_provider_check:

            If Len(m_password) > 0 Then
                m_connectiondata.PWD.Name = "PWD"
                m_connectiondata.PWD.Value = m_password
            End If

            If Len(m_username) > 0 Then
                m_connectiondata.UID.Name = "UID"
                m_connectiondata.UID.Value = m_username
            End If

            If m_ConnectionTimeout > 0 Then
                m_connectiondata.TIMEOUT.Name = "Connect Timeout"
                m_connectiondata.TIMEOUT.Value = m_ConnectionTimeout
            End If

            'override provider type (if specified)
            If providertype <> enumProvider.provider_AUTO Then m_connectiondata.ProviderType = providertype

            Select Case m_connectiondata.ProviderType
#If Support_Access Then
                Case enumProvider.provider_MSACCESS
                    buildconnection(0) = m_connectiondata.PROVIDER.Name & "=" & m_connectiondata.PROVIDER.Value & ";"
                    buildconnection(1) = m_connectiondata.DATASOURCE.Name & "=" & m_connectiondata.DATASOURCE.Value & ";"
                    If Len(m_connectiondata.EXTENDEDPROPERTIES.Value) > 0 Then buildconnection(2) = m_connectiondata.EXTENDEDPROPERTIES.Name & "=" & m_connectiondata.EXTENDEDPROPERTIES.Value & ";"
                    If m_pooling = False Then buildconnection(3) = "OLE DB Services=-2;"
                    m_connectiondata.CONNECTIONSTRING = String.Join("", buildconnection)
#End If
#If Support_MSSQL Then
                Case enumProvider.provider_SQLSERVER
                    buildconnection(1) = If(Len(m_connectiondata.NETWORK.Name) = 0, "", m_connectiondata.NETWORK.Name & "=" & m_connectiondata.NETWORK.Value & ";")
                    buildconnection(2) = m_connectiondata.SERVER.Name & "=" & m_connectiondata.SERVER.Value & ";"
                    buildconnection(3) = m_connectiondata.DATABASE.Name & "=" & m_connectiondata.DATABASE.Value & ";"
                    buildconnection(4) = m_connectiondata.UID.Name & "=" & m_connectiondata.UID.Value & ";"
                    buildconnection(5) = If(Len(m_connectiondata.PWD.Name) = 0, "", m_connectiondata.PWD.Name & "=" & m_connectiondata.PWD.Value & ";")
                    If m_connectiondata.TIMEOUT.Value > 0 Then buildconnection(7) = m_connectiondata.TIMEOUT.Name & "=" & m_connectiondata.TIMEOUT.Value & ";"
                    If m_pooling = False Then buildconnection(8) = "POOLING=False;"
                    If Len(m_application) > 0 Then
                        m_connectiondata.APPLICATIONNAME.Name = "Application Name"
                        m_connectiondata.APPLICATIONNAME.Value = m_application
                    End If
                    buildconnection(9) = If(Len(m_connectiondata.APPLICATIONNAME.Name) = 0, "", m_connectiondata.APPLICATIONNAME.Name & "=" & m_connectiondata.APPLICATIONNAME.Value & ";")
                    buildconnection(10) = If(Len(m_connectiondata.MARS.Name) = 0, "", m_connectiondata.MARS.Name & "=" & m_connectiondata.MARS.Value & ";")
                    m_connectiondata.CONNECTIONSTRING = String.Join("", buildconnection)
#End If
#If Support_MSSQL_Compact Then
                Case enumProvider.provider_SQLSERVER_COMPACT
                    buildconnection(1) = m_connectiondata.DATASOURCE.Name & "=" & m_connectiondata.DATASOURCE.Value & ";"
                    'If m_pooling = False Then buildconnection(3) = "POOLING=False;"
                    m_connectiondata.CONNECTIONSTRING = String.Join("", buildconnection)
#End If
#If Support_MySQL Then
                Case enumProvider.provider_MYSQL
                    buildconnection(1) = m_connectiondata.SERVER.Name & "=" & m_connectiondata.SERVER.Value & ";"
                    buildconnection(2) = m_connectiondata.DATABASE.Name & "=" & m_connectiondata.DATABASE.Value & ";"
                    buildconnection(3) = m_connectiondata.UID.Name & "=" & m_connectiondata.UID.Value & ";"
                    buildconnection(4) = m_connectiondata.PWD.Name & "=" & m_connectiondata.PWD.Value & ";"
                    If m_connectiondata.TIMEOUT.Value > 0 Then buildconnection(6) = m_connectiondata.TIMEOUT.Name & "=" & m_connectiondata.TIMEOUT.Value & ";"
                    m_connectiondata.CONNECTIONSTRING = String.Join("", buildconnection)
#End If
#If Support_Oracle Then
                Case enumProvider.provider_ORACLE
                    buildconnection(0) = m_connectiondata.PROVIDER.Name & "=" & m_connectiondata.PROVIDER.Value & ";"
                    buildconnection(1) = m_connectiondata.SERVER.Name & "=" & m_connectiondata.SERVER.Value & ";"
                    buildconnection(2) = m_connectiondata.DATABASE.Name & "=" & m_connectiondata.DATABASE.Value & ";"
                    buildconnection(3) = m_connectiondata.UID.Name & "=" & m_connectiondata.UID.Value & ";"
                    buildconnection(4) = m_connectiondata.PWD.Name & "=" & m_connectiondata.PWD.Value & ";"
                    If m_connectiondata.TIMEOUT.Value > 0 Then buildconnection(6) = m_connectiondata.TIMEOUT.Name & "=" & m_connectiondata.TIMEOUT.Value & ";"
                    m_connectiondata.CONNECTIONSTRING = String.Join("", buildconnection)
#End If
#If Support_SQLite Then
                Case enumProvider.provider_SQLITE
                    'buildconnection(0) = m_connectiondata.PROVIDER.Name & "=" & m_connectiondata.PROVIDER.Value & ";"
                    buildconnection(1) = m_connectiondata.DATASOURCE.Name & "=" & m_connectiondata.DATASOURCE.Value & ";"
                    'If m_pooling = False Then buildconnection(3) = "POOLING=False;"
                    m_connectiondata.CONNECTIONSTRING = String.Join("", buildconnection)
#End If
#If Support_TXT Then
                Case enumProvider.provider_TXT
                    buildconnection(0) = "DRIVER={Microsoft Text Driver (*.txt; *.csv)};"
                    If len(m_connectiondata.EXTENSIONS.Value) = 0  Then buildconnection(1) = "EXTENSIONS=asc,csv,tab,txt;" Else buildconnection(1) = m_connectiondata.EXTENSIONS.Name & "=" & m_connectiondata.EXTENSIONS.Value & ";"
                    If len(m_connectiondata.DBQ.Value) = 0  Then buildconnection(2) = "DBQ=" & m_ConnectionString & ";" Else buildconnection(2) = m_connectiondata.DBQ.Name & "=" & m_connectiondata.DBQ.Value & ";"
                    If Len(m_connectiondata.EXTENDEDPROPERTIES.Value) > 0 Then buildconnection(3) = m_connectiondata.EXTENDEDPROPERTIES.Name & "=" & m_connectiondata.EXTENDEDPROPERTIES.Value & ";"
                    m_connectiondata.CONNECTIONSTRING = String.Join("", buildconnection)
#End If
                Case Else
                    m_connectiondata.CONNECTIONSTRING = m_ConnectionString
            End Select
        End Sub

        Protected Friend Function Get_ProviderSpecific_DataAdapter() As Data.IDataAdapter
            If False Then
                Return Nothing
#If Support_MSSQL Then
            ElseIf TypeOf connection Is Data.SqlClient.SqlConnection Then
                Return New Data.SqlClient.SqlDataAdapter
#End If
#If Support_MSSQL_Compact Then
            ElseIf TypeOf connection Is Data.SqlServerCe.SqlCeConnection Then
                Return New Data.SqlServerCe.SqlCeDataAdapter
#End If
#If Support_MySQL Then
            ElseIf TypeOf connection Is MySql.Data.MySqlClient.MySqlConnection Then
                Return New MySql.Data.MySqlClient.MySqlDataAdapter
#End If
#If Support_Oracle Then
            ElseIf TypeOf connection Is Data.OracleClient.OracleConnection Then
                Return New Data.OracleClient.OracleDataAdapter
#End If
#If Support_SQLite Then
            ElseIf TypeOf connection Is Data.SQLite.SQLiteConnection Then
                Return New Data.SQLite.SQLiteDataAdapter
#End If
            Else
                Return New Data.OleDb.OleDbDataAdapter
            End If
        End Function

        Protected Friend Function Get_ProviderSpecific_Command() As Data.IDbCommand
            Dim command As Data.IDbCommand = Nothing
            If False Then
#If Support_MSSQL Then
            ElseIf TypeOf connection Is Data.SqlClient.SqlConnection Then
                command = New Data.SqlClient.SqlCommand
#End If
#If Support_MSSQL_Compact Then
            ElseIf TypeOf connection Is Data.SqlServerCe.SqlCeConnection Then
                command = New Data.SqlServerCe.SqlCeCommand
#End If
#If Support_MySQL Then
            ElseIf TypeOf connection Is MySql.Data.MySqlClient.MySqlConnection Then
                command = New MySql.Data.MySqlClient.MySqlCommand
#End If
#If Support_Oracle Then
            ElseIf TypeOf connection Is Data.OracleClient.OracleConnection Then
                command = New Data.OracleClient.OracleCommand
#End If
#If Support_SQLite Then
            ElseIf TypeOf connection Is Data.SQLite.SQLiteConnection Then
                command = New Data.SQLite.SQLiteCommand
#End If
            Else
                command = New Data.OleDb.OleDbCommand
            End If
            If Not connection_transaction Is Nothing AndAlso Not command Is Nothing Then command.Transaction = connection_transaction
            Return command
        End Function

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    If m_State <> ObjectStateEnum.adStateClosed Then Connection_Cleanup()
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Public Class Errors
        Implements GenericEnumerableReadOnly(Of [Error])
        Implements IDisposable

        Protected Friend conn As Connection

        Public ReadOnly Property Count() As Integer Implements GenericEnumerableReadOnly(Of [Error]).Count
            Get
                Return 0
            End Get
        End Property

        Public ReadOnly Property Item(ByVal index As Integer) As [Error] Implements GenericEnumerableReadOnly(Of [Error]).Item
            Get
                Return New [Error]() With {.conn = conn}
            End Get
        End Property

        Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of [Error]) Implements System.Collections.Generic.IEnumerable(Of [Error]).GetEnumerator
            Return New GenericEnumeratorReadOnly(Of [Error])(Me)
        End Function

        Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return New GenericEnumeratorReadOnly(Of [Error])(Me)
        End Function

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Public Class [Error]
        Implements IDisposable

        Protected Friend conn As Connection

        Public Shared Narrowing Operator CType(ByVal err As [Error]) As String
            Return err.Description
        End Operator

        Public ReadOnly Property Number() As Integer
            Get
                If Not conn.m_Errors Is Nothing Then Return Runtime.InteropServices.Marshal.GetHRForException(conn.m_Errors)
                Return 0
            End Get
        End Property

        Public ReadOnly Property Description() As String
            Get
                If Not conn.m_Errors Is Nothing Then Return conn.m_Errors.Message
                Return ""
            End Get
        End Property

        Public ReadOnly Property Source() As String
            Get
                If Not conn.m_Errors Is Nothing Then Return conn.m_Errors.Source
                Return ""
            End Get
        End Property

        Public ReadOnly Property SQLState() As String
            Get
                If Not conn.m_Errors Is Nothing AndAlso Not conn.m_Errors.InnerException Is Nothing Then Return conn.m_Errors.InnerException.Message
                Return ""
            End Get
        End Property

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Public Class Recordset
        Implements IDisposable

        Public data_table As Data.DataTable 'the current data page loaded up
        Protected Friend schema_table As Data.DataTable 'the current schema table
        Protected Friend schema_table_has_key_info As Boolean 'true/false if the schema table has pulled the key/unique index information (faster to not pull it unless someone requests it specifically)
        Protected Friend data_row As Data.DataRow 'the current row of data being added
        Protected Friend dictFields As Dictionary(Of String, Integer) 'used to quick lookup the field index by fieldname
        ''' <summary>
        ''' Keep track of fields that have been modified.<br />
        ''' Fields set to the same value that they already are, will NOT be added to this collection.
        ''' </summary>
        Public updatedFields As Dictionary(Of Integer, String)
        ''' <summary>
        ''' Keep track of fields that have been accessed.<br />
        ''' This is automatically collected when updating fields.<br />
        ''' To also collect this information when reading fields, set the property .EnabledAccessedFieldsWhenReading = true before accessing any fields.
        ''' </summary>
        Public accessedFields As Dictionary(Of Integer, String)
        ''' <summary>
        ''' Set this to true to have accessedFields automatically fill in when reading fields.<br />
        ''' (normally accessedFields only collects information when updating fields)
        ''' </summary>
        Public enableAccessedFieldsWhenReading As Boolean

        Protected Friend RecordsAffected As Integer

        ''' <summary>
        ''' When loading data into the DataTable, do it in such a way that constraints will not be used<br />
        ''' This is useful if trying to select and update a single column in a server side cursor that is normally part of a unique index<br />
        ''' <br />
        ''' Example:<br />
        ''' <c>DYN.Open("SELECT User_Name From Sendout_SplitDesk WHERE User_Name='FRED'", CONN, adOpenDynamic, adLockOptimistic)</c><br />
        ''' Then loop through setting <c>DYN("User_Name") = "BARNEY"</c><br />
        ''' This User_Name field is part of a unique index,<br />
        ''' so it only ever retrieves 1 record with constraints enabled unless the other fields are selected.<br />
        ''' <br />
        ''' NOTE: this setting will be ignored for client side cursors, as it would be very dangerous to disable constraints.
        ''' </summary>
        Public DisableDataTableConstraints As Boolean

        'used for Find()/Test() type functions
        Private find_rows As Data.DataRow() = Nothing
        Private find_criteria As String = ""

        'ADODB.Recordset.xxxx properties
        Protected Friend m_EditMode As EditModeEnum
        Private m_RecordCount As Integer
        Private m_CursorType As CursorTypeEnum = CursorTypeEnum.adOpenStatic 'default to static if none specified
        Private m_LockType As LockTypeEnum = LockTypeEnum.adLockReadOnly 'default if none specified
        Private m_AbsolutePosition As Integer 'current record
        Private m_AbsolutePage As Integer 'current page
        Private m_State As ObjectStateEnum = ObjectStateEnum.adStateClosed
        Private m_PageSize As Integer = 50 'default records per page
        Private m_PageCount As Integer
        Private m_EOP As Boolean 'end of page (EOF is calculated on the fly)
        Private m_Fields As Fields
        Private m_Source As String
        Private m_CursorLocation As CursorLocationEnum
        Private m_ActiveConnection As Connection

        Private data_adapter As Data.Common.DbDataAdapter

        Private last_position As Integer 'last absolute position we were on
        Private last_page As Integer 'last page we were on
        Protected Friend command As Data.IDbCommand

        Protected Friend loaded_start As Integer 'starting record loaded in data_table
        Protected Friend loaded_end As Integer 'ending record loaded in data_table
        Protected Friend page_deleted_count As Integer 'number of currently deleted items in the currently loaded page (used for server side cursor deletes)
        Protected Friend total_deleted_count As Integer 'total number of currently deleted items (used for server side cursor deletes)
        Protected Friend cursor_id As Integer 'server side cursor id (-1 = clientside)

        Public ReadOnly Property RecordCount() As Integer
            Get
                If m_State <> ObjectStateEnum.adStateOpen Then DoError(eErrorType.ERROR_CLOSED)
                If m_RecordCount = -1 Then LoadPage()
                Return m_RecordCount
            End Get
        End Property

        Public Property AbsolutePosition() As PositionEnum
            Get
                Return m_AbsolutePosition
            End Get
            Set(ByVal Value As PositionEnum)
                Select Case Value
                    Case PositionEnum.adPosBOF
                        Value = 0
                    Case PositionEnum.adPosEOF
                        Value = RecordCount + 1
                End Select
                last_position = m_AbsolutePosition
                last_page = m_AbsolutePage
                m_AbsolutePosition = Value
                m_AbsolutePage = Math.Truncate((Value - 1) / m_PageSize) + 1
            End Set
        End Property

        Public Property CacheSize() As Integer 'I am using PageSize the same as CacheSize, so this code is identical
            Get
                Return PageSize
            End Get
            Set(ByVal value As Integer)
                PageSize = value
            End Set
        End Property

        Public Property PageSize() As Integer
            Get
                Return m_PageSize
            End Get
            Set(ByVal value As Integer)
                If value > 0 Then
                    m_PageSize = value
                    AbsolutePosition = AbsolutePosition
                    m_PageCount = Math.Truncate((m_RecordCount - 1) / m_PageSize) + 1
                End If
            End Set
        End Property

        Public ReadOnly Property PageCount() As Integer
            Get
                Return m_PageCount
            End Get
        End Property

        Public Property AbsolutePage() As Integer
            Get
                Return m_AbsolutePage
            End Get
            Set(ByVal value As Integer)
                last_position = m_AbsolutePosition
                last_page = m_AbsolutePage
                m_AbsolutePage = value
                m_AbsolutePosition = (value - 1) * m_PageSize + 1
            End Set
        End Property

        Public ReadOnly Property EOF() As Boolean
            Get
                If m_State <> ObjectStateEnum.adStateOpen Then DoError(eErrorType.ERROR_CLOSED)
                Return (AbsolutePosition > RecordCount OrElse RecordCount = 0) AndAlso m_EditMode <> EditModeEnum.adEditAdd
            End Get
        End Property

        Public ReadOnly Property EOP() As Boolean 'end of page
            Get
                If m_State <> ObjectStateEnum.adStateOpen Then DoError(eErrorType.ERROR_CLOSED)
                Return m_EOP
            End Get
        End Property

        Public ReadOnly Property BOF() As Boolean
            Get
                If m_State <> ObjectStateEnum.adStateOpen Then DoError(eErrorType.ERROR_CLOSED)
                Return (AbsolutePosition < 1 OrElse RecordCount = 0) AndAlso m_EditMode <> EditModeEnum.adEditAdd
            End Get
        End Property

        Public Sub Close()
            If m_State = ObjectStateEnum.adStateClosed Then DoError(eErrorType.ERROR_CLOSED)
            If m_EditMode <> EditModeEnum.adEditNone Then DoError(eErrorType.ERROR_CONTEXT)

            Recordset_Cleanup()

            'because this object can be reused, reinitialize approriate variables here
            last_page = 0
            last_position = 0
            loaded_start = 0
            loaded_end = 0
            find_rows = Nothing
            find_criteria = ""

            m_Source = ""
            m_RecordCount = -1
            m_PageCount = -1
            m_AbsolutePosition = 0
            'm_PageSize = 10 'the COM ADODB does not reset the pagesize when closed
            m_AbsolutePage = 1
            m_State = ObjectStateEnum.adStateClosed
        End Sub

        Public Property CursorLocation() As CursorLocationEnum
            Get
                If m_CursorLocation = 0 Then Return CursorLocationEnum.adUseServer 'return default of serverside if not set
                Return m_CursorLocation
            End Get
            Set(ByVal Value As CursorLocationEnum)
                If m_State <> ObjectStateEnum.adStateClosed Then DoError(eErrorType.ERROR_OPEN)
                m_CursorLocation = Value
            End Set
        End Property

        Public Property CursorType() As CursorTypeEnum
            Get
                Return m_CursorType
            End Get
            Set(ByVal Value As CursorTypeEnum)
                m_CursorType = Value
            End Set
        End Property

        Public Property LockType() As LockTypeEnum
            Get
                Return m_LockType
            End Get
            Set(ByVal Value As LockTypeEnum)
                m_LockType = Value
            End Set
        End Property

        Public Property ActiveConnection() As Connection
            Get
                Return m_ActiveConnection
            End Get
            Set(ByVal Value As Connection)
                If m_State <> ObjectStateEnum.adStateClosed Then DoError(eErrorType.ERROR_OPEN)
                m_ActiveConnection = Value
            End Set
        End Property

        Public ReadOnly Property Fields() As Fields
            Get
                Return m_Fields
            End Get
        End Property

        Public Property Fields(ByVal index As Integer) As Field
            Get
                Return m_Fields(index)
            End Get
            Set(ByVal Value As Field)
                m_Fields(index).Value = Value.Value
            End Set
        End Property

        Public Property Fields(ByVal index As String) As Field
            Get
                Return m_Fields(index)
            End Get
            Set(ByVal Value As Field)
                m_Fields(index).Value = Value.Value
            End Set
        End Property

        Default Public Property FieldObject(ByVal index As Integer) As Field
            Get
                Return m_Fields(index)
            End Get
            Set(ByVal Value As Field)
                m_Fields(index).Value = Value.Value
            End Set
        End Property

        Default Public Property FieldObject(ByVal index As String) As Field
            Get
                Return m_Fields(index)
            End Get
            Set(ByVal Value As Field)
                m_Fields(index).Value = Value.Value
            End Set
        End Property

        Public Function GetString(Optional ByVal StringFormat As StringFormatEnum = StringFormatEnum.adClipString, Optional ByVal NumRows As Integer = -1, Optional ByVal ColumnDelimiter As String = vbTab, Optional ByVal RowDelimiter As String = vbCr, Optional ByVal NullExpr As String = "") As String
            Dim rows As New List(Of String)
            Dim columns As New List(Of String)
            Dim columnindex As Integer
            Dim rowindex As Integer

            If NumRows = 0 Then DoError(eErrorType.ERROR_CONTEXT)
            If BOF OrElse EOF Then DoError(eErrorType.ERROR_OUT_OF_BOUNDS)

            rowindex = 0
            Do While Not EOF
                If rowindex = NumRows Then Exit Do
                rowindex += 1

                columns.Clear()
                For columnindex = 0 To m_Fields.Count - 1
                    If Fields(columnindex).Value Is DBNull.Value Then
                        columns.Add(NullExpr)
                    Else
                        columns.Add(Fields(columnindex).Value.ToString())
                    End If
                Next
                rows.Add(String.Join(ColumnDelimiter, columns))
                MoveNext()
            Loop
            rows.Add("") 'make sure there is a RowDelimiter at the end
            Return String.Join(RowDelimiter, rows)
        End Function

        ''' <summary>
        ''' Run a "Select" on the current row and return True/False if it matches.
        ''' Note: this will see changes if you are in the middle of an edit.
        ''' </summary>
        ''' <param name="Criteria">WHERE test criteria (i.e. Specialty_2='C')</param>
        ''' <returns></returns>
        Public Function TestCurrentRow(Criteria As String) As Boolean
            'sanity checks
            If State = ObjectStateEnum.adStateClosed Then Return False
            If LCase(Criteria) = "true" Then Return True

            'to make this work we need to create a "dummy" DataTable that we can use to test just 1 row
            Static m_testTable As Data.DataTable = Nothing 'used for testing current row (similar to Find(), but only tests the current row returning true/false)

            If m_testTable Is Nothing OrElse Len(m_testTable.TableName) = 0 OrElse m_testTable.TableName <> data_table.TableName OrElse m_testTable.Columns.Count <> data_table.Columns.Count Then
                If Not m_testTable Is Nothing Then m_testTable.Dispose()

                'since we are only testing 1 row, it twice as fast to create a blank DataTable and only add columns
                'as opposed to cloning the existing DataTable, which also copies the schema, constraints, and other things we don't need to run a simple test
                m_testTable = New Data.DataTable(data_table.TableName)
                For i = 0 To data_table.Columns.Count - 1
                    Dim c = data_table.Columns.Item(i)
                    m_testTable.Columns.Add(c.ColumnName, c.DataType)
                Next
            End If

            'load up the current row into our temp table
            Dim currow = GetCurrentRow()
            m_testTable.LoadDataRow(currow.ItemArray(), loadOption:=Data.LoadOption.PreserveChanges)

            Try
                Dim testrow = m_testTable.Select(Criteria) 'run test on the one row
                Return If(testrow.Length > 0, True, False) 'if the Select() returned one row, the test matched, return true
            Catch ex As Exception
                'an error will be thrown if a column is missing from the select criteria, since we can't check it log it and return false for now
                'Debug_Log(ex)
                Return False
            Finally
                If Not m_testTable Is Nothing Then m_testTable.Rows.Clear()
            End Try
        End Function

        ''' <summary>
        ''' Jump to the next matching row based on the WHERE criteria.  If you keep calling this with the same criteria, it will continue to the next record.
        ''' </summary>
        ''' <param name="Criteria">WHERE test criteria (i.e. Specialty_2='C')</param>
        ''' <param name="SkipRecords"></param>
        ''' <param name="SearchDirection"></param>
        Public Sub Find(ByVal Criteria As String, Optional ByVal SkipRecords As Integer = 0, Optional ByVal SearchDirection As SearchDirectionEnum = SearchDirectionEnum.adSearchForward)
            If Criteria <> find_criteria Then
                find_rows = Nothing
                find_criteria = Criteria
            End If

            For i = 1 To SkipRecords
                If BOF OrElse EOF Then Exit For
                Select Case SearchDirection
                    Case SearchDirectionEnum.adSearchForward
                        MoveNext()
                    Case SearchDirectionEnum.adSearchBackward
                        MovePrevious()
                End Select
            Next

            Do
                If Not find_rows Is Nothing Then
                    Do While Not BOF AndAlso Not EOF
                        If CheckPage() Then
                            'if a new page has been loaded, we need to run DataTable.Select() again
                            find_rows = Nothing
                            Exit Do
                        End If

                        For i As Integer = 0 To find_rows.Length - 1
                            If data_table.Rows(AbsolutePosition - loaded_start).Equals(find_rows(i)) Then Exit Do 'found the matching row
                        Next

                        'didn't find a match for the current record, move to the next record
                        Select Case SearchDirection
                            Case SearchDirectionEnum.adSearchForward
                                MoveNext()
                            Case SearchDirectionEnum.adSearchBackward
                                MovePrevious()
                        End Select
                    Loop

                    If Not find_rows Is Nothing Then
                        'good match, return
                        Return
                    End If
                End If

                If find_rows Is Nothing Then
                    Do While Not BOF AndAlso Not EOF
                        find_rows = data_table.Select(Criteria)
                        If find_rows.Length > 0 Then Exit Do
                        Select Case SearchDirection
                            Case SearchDirectionEnum.adSearchForward
                                Dim movecount As Integer = loaded_end - AbsolutePosition
                                If AbsolutePosition + movecount > RecordCount Then movecount = RecordCount - AbsolutePosition
                                Move(movecount + 1)
                                CheckPage()
                            Case SearchDirectionEnum.adSearchBackward
                                Dim movecount As Integer = loaded_start - AbsolutePosition
                                If AbsolutePosition + movecount < 1 Then movecount = -AbsolutePosition
                                Move(movecount - 1)
                                CheckPage()
                        End Select
                    Loop
                    If BOF OrElse EOF Then Return
                End If
            Loop
        End Sub

        Public Sub Move(ByVal NumRecords As Integer, Optional ByVal Start As Integer = -1)
            If m_State <> ObjectStateEnum.adStateOpen Then DoError(eErrorType.ERROR_CLOSED)
            If m_EditMode <> EditModeEnum.adEditNone Then Update()
            AbsolutePosition = IIf(Start > -1, Start, AbsolutePosition) + NumRecords
            If AbsolutePosition > RecordCount + 1 OrElse AbsolutePosition < 0 Then DoError(eErrorType.ERROR_OUT_OF_BOUNDS)
            m_EOP = IIf(AbsolutePosition < loaded_start OrElse AbsolutePosition > loaded_end, True, False)
        End Sub

        Public Sub MoveFirst()
            Move(1, 0)
        End Sub

        Public Sub MoveLast()
            Move(RecordCount, 0)
        End Sub

        Public Sub MoveNext()
            Move(1)
        End Sub

        Public Sub MovePrevious()
            Move(-1)
        End Sub

        Protected Friend Sub InitOpen(ByRef data As Data.DataTable, Optional ByVal ActiveConnection As Connection = Nothing)
            If m_State <> ObjectStateEnum.adStateClosed Then DoError(eErrorType.ERROR_OPEN)

            If Not ActiveConnection Is Nothing Then m_ActiveConnection = ActiveConnection

            data_adapter = m_ActiveConnection.Get_ProviderSpecific_DataAdapter()
            command = m_ActiveConnection.Get_ProviderSpecific_Command()
            data_table = data
            data_table.Constraints.Clear()
            schema_table = Nothing

            command.CommandTimeout = m_ActiveConnection.CommandTimeout
            command.Connection = m_ActiveConnection.connection
            data_adapter.SelectCommand = command

            'load field collection object
            m_Fields = New Fields() With {.RS = Me}

            Try
                m_ActiveConnection.ChildObjects.Add(Me)
            Catch
            End Try

            m_RecordCount = data.Rows.Count
            m_PageCount = m_RecordCount
            m_PageCount = 1
            loaded_start = If(m_RecordCount = 0, 0, 1)
            loaded_end = m_RecordCount
            cursor_id = -1
            m_EditMode = EditModeEnum.adEditNone
            AbsolutePosition = 1
            updatedFields = Nothing
            accessedFields = Nothing

            dictFields = New Dictionary(Of String, Integer)(data_table.Columns.Count, StringComparer.OrdinalIgnoreCase)

            For i As Integer = 0 To data_table.Columns.Count - 1
                Dim new_field As Field = New Field() With {.RS = Me, .column = i}
                m_Fields.Add(new_field, data_table.Columns(i).ColumnName)
            Next
            m_State = ObjectStateEnum.adStateOpen
        End Sub

        Protected Friend Sub InitOpen(ByRef command As Data.IDbCommand, Optional ByVal ActiveConnection As Connection = Nothing, Optional ByVal CursorType As CursorTypeEnum = CursorTypeEnum.adOpenUnspecified, Optional ByVal LockType As LockTypeEnum = LockTypeEnum.adLockUnspecified, Optional ByVal Options As Integer = -1)
            If m_State <> ObjectStateEnum.adStateClosed Then DoError(eErrorType.ERROR_OPEN)

            If Not ActiveConnection Is Nothing Then
                m_ActiveConnection = ActiveConnection
                If m_CursorLocation = 0 Then m_CursorLocation = m_ActiveConnection.CursorLocation 'set the cursor location to match the connection if it has never been set
            End If
            If CursorType <> CursorTypeEnum.adOpenUnspecified Then m_CursorType = CursorType
            If LockType <> LockTypeEnum.adLockUnspecified Then m_LockType = LockType

            data_adapter = m_ActiveConnection.Get_ProviderSpecific_DataAdapter()
            If command Is Nothing Then command = m_ActiveConnection.Get_ProviderSpecific_Command()
            data_table = New Data.DataTable()
            schema_table = Nothing

            command.CommandTimeout = m_ActiveConnection.CommandTimeout
            command.Connection = m_ActiveConnection.connection
            data_adapter.SelectCommand = command

            'load field collection object
            m_Fields = New Fields() With {.RS = Me}

            Try
                m_ActiveConnection.ChildObjects.Add(Me)
            Catch
            End Try

            m_RecordCount = -1
            m_PageCount = -1
            cursor_id = -1
            AbsolutePosition = 0
            page_deleted_count = 0
            total_deleted_count = 0
            m_EditMode = EditModeEnum.adEditNone
            updatedFields = Nothing
            accessedFields = Nothing
        End Sub

        ''' <summary>
        ''' </summary>
        ''' <param name="Source"></param>
        ''' <param name="ActiveConnection"></param>
        ''' <param name="CursorType"></param>
        ''' <param name="LockType"></param>
        ''' <param name="Options">Not implemented</param>
        Public Sub Open(ByVal Source As Command, Optional ByVal ActiveConnection As Connection = Nothing, Optional ByVal CursorType As CursorTypeEnum = CursorTypeEnum.adOpenUnspecified, Optional ByVal LockType As LockTypeEnum = LockTypeEnum.adLockUnspecified, Optional ByVal Options As Integer = -1)
            command = Source
            InitOpen(command, ActiveConnection, CursorType, LockType, Options)
            LoadPage() 'load the first page of data
            m_State = ObjectStateEnum.adStateOpen
        End Sub

        ''' <summary>
        ''' </summary>
        ''' <param name="Source"></param>
        ''' <param name="ActiveConnection"></param>
        ''' <param name="CursorType"></param>
        ''' <param name="LockType"></param>
        ''' <param name="Options">Not implemented</param>
        Public Sub Open(Optional ByVal Source As String = "", Optional ByVal ActiveConnection As Connection = Nothing, Optional ByVal CursorType As CursorTypeEnum = CursorTypeEnum.adOpenUnspecified, Optional ByVal LockType As LockTypeEnum = LockTypeEnum.adLockUnspecified, Optional ByVal Options As Integer = -1)
            InitOpen(command, ActiveConnection, CursorType, LockType, Options)
            If Len(Source) > 0 Then m_Source = Source
            LoadPage() 'load the first page of data
            m_State = ObjectStateEnum.adStateOpen
        End Sub

        'same as Open(string), but separates out the parameters (similar to a parameterized query) to centralize the SQL injection protection code
        'this method will also allow a server-side "parameterized" query, which conventional parameterization makes this impossible to do
        Public Sub OpenEx(ByVal Source As String, ByVal ActiveConnection As Connection, ByVal CursorType As CursorTypeEnum, ByVal LockType As LockTypeEnum, ByVal ParamArray parameters() As Object)
            InitOpen(command, ActiveConnection, CursorType, LockType)

            'replace parameters in query
            If UBound(parameters) > -1 Then
                Dim new_source As New System.Text.StringBuilder(Len(Source))
                Dim parm_pos As Integer = 0
                For i As Integer = 0 To UBound(parameters)
                    Dim pos = InStr(parm_pos + 1, Source, "?")
                    If pos = 0 Then Exit For
                    new_source.Append(Mid(Source, parm_pos + 1, pos - parm_pos - 1))
                    parm_pos = pos
                    If TypeOf parameters(i) Is Array Then
                        If TypeOf parameters(i) Is Char() Then
                            'char array
                            GoTo do_string
                        ElseIf TypeOf parameters(i) Is String() Then
                            'string array
                            new_source.Append(CType(parameters(i), String()).SQLNest())
                        ElseIf TypeOf parameters(i) Is Int32() Then
                            'integer array
                            new_source.Append(CType(parameters(i), Int32()).SQLNest())
                        Else
                            'TODO: support other array types
                        End If
                    Else
                        If TypeOf parameters(i) Is Guid Then
                            'guid
                            new_source.Append("'")
                            new_source.Append(Replace(CType(parameters(i), Guid).ToString(), "'", "''")) 'TODO: MySQL also has to take backslashes into account
                            new_source.Append("'")
                        ElseIf TypeOf parameters(i) Is Date Then
                            'date
                            Dim date_separator As String
                            If TypeOf ActiveConnection.connection Is Data.OleDb.OleDbConnection Then date_separator = "#" Else date_separator = "'"
                            new_source.Append(date_separator)
                            new_source.Append(Replace(CType(parameters(i), Date).ToString("yyyy-MM-dd HH:mm:ss"), date_separator, ""))
                            new_source.Append(date_separator)
                        ElseIf TypeOf parameters(i) Is Byte Then
                            'byte (numeric for now)
                            new_source.Append(CType(parameters(i), Byte).ToString())
                        ElseIf TypeOf parameters(i) Is Double OrElse TypeOf parameters(i) Is Int32 OrElse TypeOf parameters(i) Is Int16 OrElse TypeOf parameters(i) Is Int64 OrElse TypeOf parameters(i) Is Single OrElse TypeOf parameters(i) Is Decimal Then
                            'numeric
                            new_source.Append(parameters(i))
                        ElseIf parameters(i) Is Nothing Then
                            'set to null
                            new_source.Append("null")
                        Else
do_string:
                            'string
                            new_source.Append("'")
                            new_source.Append(Replace(parameters(i).ToString(), "'", "''")) 'TODO: MySQL also has to take backslashes into account
                            new_source.Append("'")
                        End If
                    End If
                Next
                new_source.Append(Mid(Source, parm_pos + 1))
                Source = new_source.ToString()
            End If

            m_Source = Source
            LoadPage() 'load the first page of data
            m_State = ObjectStateEnum.adStateOpen
        End Sub

        Public Sub Requery(Optional ByVal Options As Integer = -1)
            If m_State = ObjectStateEnum.adStateClosed Then DoError(eErrorType.ERROR_CLOSED)
            Dim hold_position As Integer = AbsolutePosition
            Dim hold_ActiveConnection = Me.m_ActiveConnection
            Recordset_Cleanup()
            m_State = ObjectStateEnum.adStateClosed
            Open(Options:=Options, ActiveConnection:=hold_ActiveConnection)
            AbsolutePosition = hold_position
        End Sub

        Public Property Source() As String
            Get
                Return m_Source
            End Get
            Set(ByVal Value As String)
                m_Source = Value
            End Set
        End Property

        Public ReadOnly Property State() As ObjectStateEnum
            Get
                Return m_State
            End Get
        End Property

        Public ReadOnly Property EditMode() As EditModeEnum
            Get
                Return m_EditMode
            End Get
        End Property

        Public Sub AddNew()
            If m_State = ObjectStateEnum.adStateClosed Then DoError(eErrorType.ERROR_CLOSED)
            If m_EditMode <> EditModeEnum.adEditAdd Then
                data_row = data_table.NewRow()
                data_row.BeginEdit()
                updatedFields = New Dictionary(Of Integer, String)()
                accessedFields = New Dictionary(Of Integer, String)()
                If cursor_id <> -1 Then
                    'make sure all columns allow nulls so that default values/incremental fields will work, otherwise the data_table won't allow us to even insert the record
                    data_table.Constraints.Clear()
                    For Each column As Data.DataColumn In data_table.Columns
                        column.AllowDBNull = True
                    Next
                End If
                m_EditMode = EditModeEnum.adEditAdd
            End If
        End Sub

        'we do not support the AffectEnum argument at this time
        Public Sub Delete()
            If m_EditMode = EditModeEnum.adEditAdd Then
                CancelUpdate()
            Else
                m_EditMode = EditModeEnum.adEditDelete
                Update() 'ADODB calls delete and update in one function
            End If
        End Sub

        Public Sub CancelUpdate()
            Select Case m_EditMode
                Case EditModeEnum.adEditAdd
                    data_row.CancelEdit()
                    data_row = Nothing
                Case EditModeEnum.adEditInProgress
                    data_table.Rows(AbsolutePosition - loaded_start).CancelEdit()
            End Select
            m_EditMode = EditModeEnum.adEditNone
            updatedFields = Nothing
            accessedFields = Nothing
        End Sub

        Public Sub Update()
            Dim r(0) As Data.DataRow

            Select Case m_EditMode
                Case EditModeEnum.adEditAdd
                    command.Transaction = m_ActiveConnection.connection_transaction
                    r(0) = data_row
                    r(0).EndEdit()

                    data_table.Rows.Add(data_row)

                    If cursor_id = -1 Then
                        'client side cursor add

                        'Some sources don't seem to pull the schema properly, so we have to manually fill in the Insert/Update/Delete commands
                        If data_adapter.InsertCommand Is Nothing Then
                            If TypeOf data_adapter Is Data.OleDb.OleDbDataAdapter Then
                                Dim cmdbldr = New Data.OleDb.OleDbCommandBuilder(data_adapter)
                                data_adapter.InsertCommand = cmdbldr.GetInsertCommand()
                                'escape column names, because the OleDbCommandBuilder sux! (i.e. INSERT INTO (first name, last name)... --> INSERT INTO ([first name], [last name])...
                                Dim re As New Text.RegularExpressions.Regex("\([^\)]+\)")
                                Dim ct = data_adapter.InsertCommand.CommandText
                                Dim m = re.Match(ct)
                                If m.Success Then
                                    Dim orig = m.Captures.Item(0).Value
                                    orig = orig.Substring(1, orig.Length - 2)
                                    Dim fixed = "[" & orig.Replace(", ", "], [") & "]"
                                    'data_adapter.InsertCommand.CommandText = ct.Replace(orig, fixed)

                                    Dim newcmd = New Data.OleDb.OleDbCommand(ct.Replace(orig, fixed), m_ActiveConnection.connection)
                                    For Each col In data_row.ItemArray
                                        newcmd.Parameters.AddWithValue("?", col)
                                    Next
                                    data_adapter.InsertCommand = newcmd

                                End If
#If Support_MSSQL Then
                            ElseIf TypeOf data_adapter Is Data.SqlClient.SqlDataAdapter Then
                                Dim cmdbldr = New Data.SqlClient.SqlCommandBuilder(data_adapter)
                                data_adapter.InsertCommand = cmdbldr.GetInsertCommand()
#End If
#If Support_SQLite Then
                        ElseIf TypeOf data_adapter Is Data.Sqlite.SqliteDataAdapter Then
                            Dim cmdbldr = New Data.SQLite.SQLiteCommandBuilder(data_adapter)
                            data_adapter.InsertCommand = cmdbldr.GetInsertCommand()
#End If
                            End If
                        End If

                        If m_ActiveConnection.DisableUpdates Then
#If Support_MSSQL Then
                        ElseIf TypeOf data_adapter Is Data.SqlClient.SqlDataAdapter Then
                            Dim d = CType(data_adapter, Data.SqlClient.SqlDataAdapter)
                            d.Update(r)
#End If
#If Support_MSSQL_Compact Then
                        ElseIf TypeOf data_adapter Is Data.SqlServerCe.SqlCeDataAdapter Then
                            CType(data_adapter, Data.SqlServerCe.SqlCeDataAdapter).Update(r)
#End If
#If Support_MySQL Then
                        ElseIf TypeOf data_adapter Is MySql.Data.MySqlClient.MySqlDataAdapter Then
                            CType(data_adapter, MySql.Data.MySqlClient.MySqlDataAdapter).Update(r)
#End If
#If Support_Oracle Then
                        ElseIf TypeOf data_adapter Is Data.OracleClient.OracleDataAdapter Then
                            CType(data_adapter, Data.OracleClient.OracleDataAdapter).Update(r)
#End If
#If Support_SQLite Then
                        ElseIf TypeOf data_adapter Is Data.SQLite.SQLiteDataAdapter Then
                            CType(data_adapter, Data.SQLite.SQLiteDataAdapter).Update(r)
#End If
                        Else
                            CType(data_adapter, Data.OleDb.OleDbDataAdapter).Update(r)
                        End If
                    Else
                        'server side cursor update
                        If m_LockType = LockTypeEnum.adLockReadOnly Then
                            m_EditMode = EditModeEnum.adEditNone
                            DoError(eErrorType.ERROR_NOUPDATES)
                        End If

                        If Not m_ActiveConnection.DisableUpdates Then
                            command.CommandText = "sp_cursor"
                            command.CommandType = Data.CommandType.StoredProcedure
                            command.Parameters.Clear()
                            command.Parameters.Add(New Data.SqlClient.SqlParameter("cursor", cursor_id))
                            command.Parameters.Add(New Data.SqlClient.SqlParameter("optype", 4))
                            command.Parameters.Add(New Data.SqlClient.SqlParameter("rownum", 0))
                            command.Parameters.Add(New Data.SqlClient.SqlParameter("table", data_table.TableName))
                            For Each col As Integer In updatedFields.Keys
                                'make sure we use the "BaseColumnName" property for the column name, which is the REAL column name and not an alias (i.e. SELECT Col1 AS Alias1....)
                                Dim pname = Fields(col).Properties(eRECORDSET_PROPERTIES.BaseColumnName).Value.ToString()

                                If Len(pname) = 0 Then pname = updatedFields(col)
                                Dim param As Data.SqlClient.SqlParameter
                                'binary params need to have type and size defined, or the Sql Adapter throws a conversion error 
                                If Fields(col).Type = DataTypeEnum.adBinary OrElse Fields(col).Type = DataTypeEnum.adVarBinary Then
                                    param = New Data.SqlClient.SqlParameter(parameterName:=pname, dbType:=If(Fields(col).Type = DataTypeEnum.adBinary, Data.SqlDbType.Binary, Data.SqlDbType.VarBinary), size:=Fields(col).DefinedSize)
                                    param.Value = r(0).Item(col)
                                ElseIf Fields(col).Type = DataTypeEnum.adUserDefined Then
                                    'this is a user defined type, get the typename                                    
                                    param = New Data.SqlClient.SqlParameter(parameterName:=pname, value:=r(0).Item(col))
                                    param.UdtTypeName = Fields(col).UdtTypeName
                                Else
                                    param = New Data.SqlClient.SqlParameter(parameterName:=pname, value:=r(0).Item(col))
                                End If
                                command.Parameters.Add(param)
                            Next

                            Using New RunSQLCancellable(command, m_ActiveConnection.CancellationToken)
                                command.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                    If m_ActiveConnection.connection_transaction Is Nothing Then data_table.AcceptChanges()
                    m_RecordCount = data_table.Rows.Count
                    m_PageCount = Math.Truncate((m_RecordCount - 1) / m_PageSize) + 1 'leave m_ variables
                    AbsolutePosition = RecordCount
                    If loaded_start = 0 Then loaded_start = 1
                    loaded_end += 1
                    data_row = Nothing

                Case EditModeEnum.adEditInProgress
                    'Some sources don't seem to pull the schema properly, so we have to manually fill in the Insert/Update/Delete commands
                    If data_adapter.UpdateCommand Is Nothing Then
                        If TypeOf data_adapter Is Data.OleDb.OleDbDataAdapter Then
                            Dim cmdbldr = New Data.OleDb.OleDbCommandBuilder(data_adapter)
                            data_adapter.UpdateCommand = cmdbldr.GetUpdateCommand()
#If Support_SQLite Then
                        ElseIf TypeOf data_adapter Is Data.SQLite.SQLiteDataAdapter Then
                            Dim cmdbldr = New Data.SQLite.SQLiteCommandBuilder(data_adapter)
                            data_adapter.UpdateCommand = cmdbldr.GetUpdateCommand()
#End If
                        End If
                    End If

                    command.Transaction = m_ActiveConnection.connection_transaction
                    r(0) = data_table.Rows(AbsolutePosition - loaded_start)
                    r(0).EndEdit()
                    If cursor_id = -1 Then
                        'client side cursor update
                        If m_ActiveConnection.DisableUpdates Then
#If Support_MSSQL Then
                        ElseIf TypeOf data_adapter Is Data.SqlClient.SqlDataAdapter Then
                            CType(data_adapter, Data.SqlClient.SqlDataAdapter).Update(r)
#End If
#If Support_MSSQL_Compact Then
                        ElseIf TypeOf data_adapter Is Data.SqlServerCe.SqlCeDataAdapter Then
                            CType(data_adapter, Data.SqlServerCe.SqlCeDataAdapter).Update(r)
#End If
#If Support_MySQL Then
                        ElseIf TypeOf data_adapter Is MySql.Data.MySqlClient.MySqlDataAdapter Then
                            CType(data_adapter, MySql.Data.MySqlClient.MySqlDataAdapter).Update(r)
#End If
#If Support_Oracle Then
                        ElseIf TypeOf data_adapter Is Data.OracleClient.OracleDataAdapter Then
                            CType(data_adapter, Data.OracleClient.OracleDataAdapter).Update(r)
#End If
#If Support_SQLite Then
                        ElseIf TypeOf data_adapter Is Data.SQLite.SQLiteDataAdapter Then
                            CType(data_adapter, Data.SQLite.SQLiteDataAdapter).Update(r)
#End If
                        Else
                            CType(data_adapter, Data.OleDb.OleDbDataAdapter).Update(r)
                        End If
                    Else
                        'server side cursor update
                        If m_LockType = LockTypeEnum.adLockReadOnly Then
                            m_EditMode = EditModeEnum.adEditNone
                            DoError(eErrorType.ERROR_NOUPDATES)
                        End If

                        If Not m_ActiveConnection.DisableUpdates AndAlso updatedFields.Count > 0 Then 'skip the update if none of the fields were updated
                            command.CommandText = "sp_cursor"
                            command.CommandType = Data.CommandType.StoredProcedure
                            command.Parameters.Clear()
                            command.Parameters.Add(New Data.SqlClient.SqlParameter("cursor", cursor_id))
                            command.Parameters.Add(New Data.SqlClient.SqlParameter("optype", 33))
                            command.Parameters.Add(New Data.SqlClient.SqlParameter("rownum", AbsolutePosition - loaded_start + page_deleted_count + 1))
                            command.Parameters.Add(New Data.SqlClient.SqlParameter("table", data_table.TableName))
                            For Each col As Integer In updatedFields.Keys
                                'make sure we use the "BaseColumnName" property for the column name, which is the REAL column name and not an alias (i.e. SELECT Col1 AS Alias1....)
                                Dim pname = Fields(col).Properties(eRECORDSET_PROPERTIES.BaseColumnName).Value.ToString()
                                If Len(pname) = 0 Then pname = updatedFields(col)

                                Dim param As Data.SqlClient.SqlParameter
                                'binary params need to have type and size defined, or the Sql Adapter throws a conversion error 
                                If Fields(col).Type = DataTypeEnum.adBinary OrElse Fields(col).Type = DataTypeEnum.adVarBinary Then
                                    param = New Data.SqlClient.SqlParameter(parameterName:=pname, dbType:=If(Fields(col).Type = DataTypeEnum.adBinary, Data.SqlDbType.Binary, Data.SqlDbType.VarBinary), size:=Fields(col).DefinedSize)
                                    param.Value = r(0).Item(col)
                                Else
                                    param = New Data.SqlClient.SqlParameter(parameterName:=pname, value:=r(0).Item(col))
                                End If
                                command.Parameters.Add(param)
                            Next

                            Using New RunSQLCancellable(command, m_ActiveConnection.CancellationToken)
                                command.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                    If m_ActiveConnection.connection_transaction Is Nothing Then data_table.AcceptChanges()

                Case EditModeEnum.adEditDelete
                    'Some sources don't seem to pull the schema properly, so we have to manually fill in the Insert/Update/Delete commands
#If Support_SQLite Then
                    If data_adapter.DeleteCommand Is Nothing Then
                        If TypeOf data_adapter Is Data.SQLite.SQLiteDataAdapter Then
                            Dim cmdbldr = New Data.SQLite.SQLiteCommandBuilder(data_adapter)
                            data_adapter.DeleteCommand = cmdbldr.GetDeleteCommand()
                        End If
                    End If
#End If

                    command.Transaction = m_ActiveConnection.connection_transaction
                    r(0) = data_table.Rows(AbsolutePosition - loaded_start)
                    r(0).BeginEdit()
                    r(0).Delete()
                    r(0).EndEdit()
                    If cursor_id = -1 Then
                        'client side cursor delete
                        If m_ActiveConnection.DisableUpdates Then
#If Support_MSSQL Then
                        ElseIf TypeOf data_adapter Is Data.SqlClient.SqlDataAdapter Then
                            CType(data_adapter, Data.SqlClient.SqlDataAdapter).Update(r)
#End If
#If Support_MSSQL_Compact Then
                        ElseIf TypeOf data_adapter Is Data.SqlServerCe.SqlCeDataAdapter Then
                            CType(data_adapter, Data.SqlServerCe.SqlCeDataAdapter).Update(r)
#End If
#If Support_MySQL Then
                        ElseIf TypeOf data_adapter Is MySql.Data.MySqlClient.MySqlDataAdapter Then
                            CType(data_adapter, MySql.Data.MySqlClient.MySqlDataAdapter).Update(r)
#End If
#If Support_Oracle Then
                        ElseIf TypeOf data_adapter Is Data.OracleClient.OracleDataAdapter Then
                            CType(data_adapter, Data.OracleClient.OracleDataAdapter).Update(r)
#End If
#If Support_SQLite Then
                        ElseIf TypeOf data_adapter Is Data.SQLite.SQLiteDataAdapter Then
                            CType(data_adapter, Data.SQLite.SQLiteDataAdapter).Update(r)
#End If
                        Else
                            CType(data_adapter, Data.OleDb.OleDbDataAdapter).Update(r)
                        End If
                    Else
                        'server side cursor delete
                        If m_LockType = LockTypeEnum.adLockReadOnly Then
                            m_EditMode = EditModeEnum.adEditNone
                            DoError(eErrorType.ERROR_NOUPDATES)
                        End If

                        If Not m_ActiveConnection.DisableUpdates Then
                            command.CommandText = "sp_cursor"
                            command.CommandType = Data.CommandType.StoredProcedure
                            command.Parameters.Clear()
                            command.Parameters.Add(New Data.SqlClient.SqlParameter("cursor", cursor_id))
                            command.Parameters.Add(New Data.SqlClient.SqlParameter("optype", 34))
                            command.Parameters.Add(New Data.SqlClient.SqlParameter("rownum", AbsolutePosition - loaded_start + page_deleted_count + 1))

                            Using New RunSQLCancellable(command, m_ActiveConnection.CancellationToken)
                                command.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                    If m_ActiveConnection.connection_transaction Is Nothing Then data_table.AcceptChanges()
                    m_RecordCount -= 1
                    m_AbsolutePosition -= 1
                    loaded_end -= 1
                    page_deleted_count += 1
                    total_deleted_count += 1
                    m_PageCount = Math.Truncate((m_RecordCount - 1) / m_PageSize) + 1 'leave m_ variables
            End Select
            m_EditMode = EditModeEnum.adEditNone
            updatedFields = Nothing
            accessedFields = Nothing
        End Sub

        Protected Friend Function CheckPage() As Boolean
            'DO NOT load a new page if we are currently adding a record, otherwise it will blast it out of existance
            If m_EditMode <> EditModeEnum.adEditAdd Then
                If loaded_start > AbsolutePosition OrElse loaded_end < AbsolutePosition OrElse loaded_end < loaded_start Then
                    LoadPage()
                    Return True
                End If
            End If
            Return False
        End Function

        Protected Friend Sub LoadPage(Optional ByVal useCommand As Data.IDbCommand = Nothing)
            If m_EditMode = EditModeEnum.adEditAdd Then Exit Sub 'SAFETY NET - DO NOT load a new page if we are currently adding a record, otherwise it will blast it out of existance

            'Dim conn As Data.IDbConnection
            'm_State = ObjectStateEnum.adStateFetching
            'conn = m_ActiveConnection.oledb_connection

            'if a command object is passed in, skip setting up the command object because it should be already done by the caller, technically this could be another function to prep the command object
            If useCommand Is Nothing Then
                'use this recordset's command object if one is not passed in
                useCommand = command

                Dim isSelect As Boolean = Left(m_Source, 6).ToUpper() = "SELECT"

#If Support_MSSQL Then
                If m_CursorLocation = CursorLocationEnum.adUseServer AndAlso (m_LockType <> LockTypeEnum.adLockReadOnly OrElse m_CursorType <> CursorTypeEnum.adOpenStatic) AndAlso TypeOf data_adapter Is Data.SqlClient.SqlDataAdapter Then
                    If isSelect Then
                        Static cursor_arg As Integer
                        Static lock_arg As Integer
                        If cursor_id = -1 Then
                            Select Case m_LockType
                                Case LockTypeEnum.adLockOptimistic, LockTypeEnum.adLockBatchOptimistic
                                    cursor_arg = &H1 'keyset cursor
                                    lock_arg = &H4004 'optimistic lock with in place update
                                Case LockTypeEnum.adLockPessimistic
                                    cursor_arg = &H1 'keyset cursor
                                    lock_arg = &H4002 'pessimistic lock with in place update
                                Case LockTypeEnum.adLockReadOnly, LockTypeEnum.adLockUnspecified
                                    Select Case m_CursorType
                                        Case CursorTypeEnum.adOpenForwardOnly, CursorTypeEnum.adOpenUnspecified
                                            GoTo client_side 'special exception, unknown will run as a clientside cursor

                                            'We are not supporting ForwardOnly server side cursors at the moment, the speed is so slow and accidental usage too high
                                            'If you absolutely must have a read-only server side cursor, use a Keyset, it's not that much slower
                                            'Case CursorTypeEnum.adOpenForwardOnly
                                            '    cursor_arg = &H4 'fast_forward
                                            'lock_arg = &H1 'readonly lock
                                        Case CursorTypeEnum.adOpenDynamic, CursorTypeEnum.adOpenKeyset
                                            cursor_arg = &H1 'keyset cursor
                                            lock_arg = &H1 'readonly lock
                                        Case CursorTypeEnum.adOpenStatic
                                            'this will never happen, because the IF clause above forces ReadOnly/Static to run client side
                                            cursor_arg = &H8 'static cursor
                                            lock_arg = &H1 'readonly lock
                                    End Select
                            End Select
                            useCommand.CommandText = "sp_cursoropen"
                            useCommand.CommandType = Data.CommandType.StoredProcedure
                            useCommand.Parameters.Clear()
                            Dim p = New Data.SqlClient.SqlParameter()
                            p.ParameterName = "cursor"
                            p.Direction = Data.ParameterDirection.Output
                            p.Value = 0
                            useCommand.Parameters.Add(p)
                            useCommand.Parameters.Add(New Data.SqlClient.SqlParameter("stmt", m_Source))
                            p = New Data.SqlClient.SqlParameter()
                            p.ParameterName = "scrollopt"
                            p.Direction = Data.ParameterDirection.InputOutput
                            p.Value = cursor_arg
                            useCommand.Parameters.Add(p)
                            p = New Data.SqlClient.SqlParameter()
                            p.ParameterName = "ccopt"
                            p.Direction = Data.ParameterDirection.InputOutput
                            p.Value = lock_arg
                            useCommand.Parameters.Add(p)
                            p = New Data.SqlClient.SqlParameter()
                            p.ParameterName = "rowcount"
                            p.Direction = Data.ParameterDirection.Output
                            p.Value = lock_arg
                            useCommand.Parameters.Add(p)

                            Using New RunSQLCancellable(useCommand, m_ActiveConnection.CancellationToken)
                                useCommand.ExecuteNonQuery()
                            End Using

                            cursor_id = CType(useCommand.Parameters("cursor"), Data.SqlClient.SqlParameter).Value
                            m_RecordCount = CType(useCommand.Parameters("rowcount"), Data.SqlClient.SqlParameter).Value
                            m_PageCount = Math.Truncate((m_RecordCount - 1) / m_PageSize) + 1 'leave m_ variables
                            If m_RecordCount > 0 AndAlso AbsolutePosition = 0 Then AbsolutePosition = 1
                        End If
                        useCommand.CommandText = "sp_cursorfetch"
                        useCommand.CommandType = Data.CommandType.StoredProcedure
                        useCommand.Parameters.Clear()
                        useCommand.Parameters.Add(New Data.SqlClient.SqlParameter("cursor", cursor_id))
                        Select Case cursor_arg
                            Case &H4, &H10 'fast_forward/forward_only cursors do not support absolute positioning
                                useCommand.Parameters.Add(New Data.SqlClient.SqlParameter("fetchtype", 2)) '2=next
                            Case Else
                                useCommand.Parameters.Add(New Data.SqlClient.SqlParameter("fetchtype", 16)) '16=absolute
                                useCommand.Parameters.Add(New Data.SqlClient.SqlParameter("rownum", AbsolutePosition + total_deleted_count))
                        End Select
                        useCommand.Parameters.Add(New Data.SqlClient.SqlParameter("nrows", If(PageSize > 0, PageSize, 1)))
                    End If
                End If
#End If

#If Support_MSSQL_Compact Then
                If m_CursorLocation = CursorLocationEnum.adUseServer AndAlso TypeOf data_adapter Is Data.SqlServerCe.SqlCeDataAdapter Then
                    If Left(m_Source, 6).ToUpper() = "SELECT" Then
                        If cursor_id = -1 Then
                            Dim cursor_arg As Integer
                            Dim lock_arg As Integer
                            Select Case m_LockType
                                Case LockTypeEnum.adLockOptimistic, LockTypeEnum.adLockBatchOptimistic
                                    cursor_arg = &H1 'keyset cursor
                                    lock_arg = &H4004 'optimistic lock with in place update
                                Case LockTypeEnum.adLockPessimistic
                                    cursor_arg = &H1 'keyset cursor
                                    lock_arg = &H4002 'pessimistic lock with in place update
                                Case LockTypeEnum.adLockReadOnly, LockTypeEnum.adLockUnspecified
                                    Select Case m_CursorType
                                        Case CursorTypeEnum.adOpenForwardOnly, CursorTypeEnum.adOpenUnspecified
                                            GoTo client_side 'special exception, forwardonly/readonly actually runs as a clientside cursor
                                        Case CursorTypeEnum.adOpenDynamic, CursorTypeEnum.adOpenKeyset
                                            cursor_arg = &H1 'keyset cursor
                                            lock_arg = &H1 'readonly lock
                                        Case CursorTypeEnum.adOpenStatic
                                            cursor_arg = &H8 'static cursor
                                            lock_arg = &H1 'readonly lock
                                    End Select
                            End Select
                            useCommand.CommandText = "sp_cursoropen"
                            useCommand.CommandType = Data.CommandType.StoredProcedure
                            useCommand.Parameters.Clear()
                            Dim p = New Data.SqlServerCe.SqlCeParameter()
                            p.ParameterName = "cursor"
                            p.Direction = Data.ParameterDirection.Output
                            p.Value = 0
                            useCommand.Parameters.Add(p)
                            useCommand.Parameters.Add(New Data.SqlServerCe.SqlCeParameter("stmt", m_Source))
                            p = New Data.SqlServerCe.SqlCeParameter()
                            p.ParameterName = "scrollopt"
                            p.Direction = Data.ParameterDirection.InputOutput
                            p.Value = cursor_arg
                            useCommand.Parameters.Add(p)
                            p = New Data.SqlServerCe.SqlCeParameter()
                            p.ParameterName = "ccopt"
                            p.Direction = Data.ParameterDirection.InputOutput
                            p.Value = lock_arg
                            useCommand.Parameters.Add(p)
                            p = New Data.SqlServerCe.SqlCeParameter()
                            p.ParameterName = "rowcount"
                            p.Direction = Data.ParameterDirection.Output
                            p.Value = lock_arg
                            useCommand.Parameters.Add(p)

                            Using New RunSQLCancellable(useCommand, m_ActiveConnection.CancellationToken)
                                useCommand.ExecuteNonQuery()
                            End Using

                            cursor_id = CType(useCommand.Parameters("cursor"), Data.SqlServerCe.SqlCeParameter).Value
                            m_RecordCount = CType(useCommand.Parameters("rowcount"), Data.SqlServerCe.SqlCeParameter).Value
                            m_PageCount = Math.Truncate((m_RecordCount - 1) / m_PageSize) + 1 'leave m_ variables
                            If m_RecordCount > 0 AndAlso AbsolutePosition = 0 Then AbsolutePosition = 1
                        End If
                        If PageSize > 0 Then
                            useCommand.CommandText = "sp_cursorfetch"
                            useCommand.CommandType = Data.CommandType.StoredProcedure
                            useCommand.Parameters.Clear()
                            useCommand.Parameters.Add(New Data.SqlServerCe.SqlCeParameter("cursor", cursor_id))
                            useCommand.Parameters.Add(New Data.SqlServerCe.SqlCeParameter("fetchtype", 16))
                            useCommand.Parameters.Add(New Data.SqlServerCe.SqlCeParameter("rownum", AbsolutePosition + total_deleted_count))
                            useCommand.Parameters.Add(New Data.SqlServerCe.SqlCeParameter("nrows", PageSize))
                        End If
                    End If
                End If
#End If

                'client side cursor loader
                If cursor_id = -1 Then
client_side:
                    useCommand.CommandText = m_Source
                    useCommand.CommandType = Data.CommandType.Text
                    useCommand.Parameters.Clear()
                End If
            End If

            Dim rdr As Data.IDataReader = Nothing
            Using New RunSQLCancellable(useCommand, m_ActiveConnection.CancellationToken)
                Try
                    rdr = useCommand.ExecuteReader()
                    If schema_table Is Nothing Then
                        'do not overwrite the schema table if we've already pulled it, because it might be the version without full schema info
                        schema_table = rdr.GetSchemaTable()

                        'server side cursors always have full schema info... (but that is about 10% slower)
                        If cursor_id <> -1 Then schema_table_has_key_info = True
                    End If

                    data_table.Clear()
                    'For Each constraint In data_table.Constraints
                    '    If constraint.GetType() = GetType(Data.UniqueConstraint) Then
                    '        Dim uc As Data.UniqueConstraint = constraint
                    '        uc.Columns(0).
                    '    End If
                    'Next
                    data_table.BeginLoadData()

                    If cursor_id = -1 OrElse Not Me.DisableDataTableConstraints Then
                        'if using a client side cursor, or if constraints are not disabled,
                        'then load the data with the data_table.Load function
                        data_table.Load(rdr)
                    Else
                        'manually load schema and data so constraints can be excluded
                        '(there doesn't seem to be any way to do this within the data_table.Load() function, so we had to write a lot of code here)
                        Dim colCount = schema_table.Rows.Count
                        If data_table.Columns.Count = 0 Then
                            For Each dataRow As Data.DataRow In schema_table.Rows
                                Dim data = New Data.DataColumn()
                                data.ColumnName = dataRow("ColumnName").ToString()
                                data.DataType = Type.GetType(dataRow("DataType").ToString())
                                data.AllowDBNull = dataRow("AllowDBNull")
                                If data.DataType = GetType(String) Then data.MaxLength = dataRow("ColumnSize")
                                'data.ReadOnly = dataRow("IsReadOnly")
                                'data.AutoIncrement = dataRow("IsAutoIncrement")
                                'data.Unique = dataRow("IsUnique")
                                'data.DefaultValue = dataRow("DefaultValue") 'this doesn't exist
                                'data.ExtendedProperties = dataRow("ExtendedProperties")
                                Do
                                    Try
                                        data_table.Columns.Add(data)
                                        Exit Do
                                    Catch
                                        data.ColumnName &= "_dupe" 'modify column name and try again
                                    End Try
                                Loop
                            Next
                        End If

                        Do While rdr.Read()
                            Dim row = data_table.NewRow()
                            Dim columns = New Object(colCount - 1) {}
                            For i = 0 To colCount - 1
                                columns(i) = rdr.GetValue(i)
                            Next
                            row.ItemArray = columns
                            data_table.Rows.Add(row)
                        Loop
                    End If
                Catch ex As Data.SqlClient.SqlException
                    Select Case ex.Number
                        Case 2601, 2627 'PRIMARY KEY/INDEX VIOLATION
                            'no logging
                        Case Else
                            'Debug_Log("MALFORMED QUERY: " & m_Source & vbCrLf & ex.StackTrace, Filename:="sqldebug")
                    End Select

                    'if you passed in a malformed query, we need to clean up or else the connection will be stuck part way through an atomic operation, and without MARS turned on, it will be dead
                    data_table.EndLoadData()
                    data_table.Constraints.Clear()
                    If Not rdr Is Nothing Then rdr.Close()
                    If Not schema_table Is Nothing Then
                        schema_table.Dispose()
                        schema_table = Nothing
                    End If
                    Throw 'rethrow
                Catch ex As Data.ConstraintException
                    'this will only happen if a SQL table has become corrupt and the data does not match the constraints
                    Dim errs = data_table.GetErrors()
                    'If UBound(errs) > -1 Then Debug_Log("ADO.NET CONSTRAINT ERROR: " & errs(0).RowError & vbCrLf & m_Source, Filename:="sqldebug")
                End Try
            End Using
            data_table.EndLoadData()
            data_table.Constraints.Clear()
            RecordsAffected = rdr.RecordsAffected
            rdr.Close()

            'set start/end markers that tells which records are currently loaded
            If cursor_id = -1 Then
                'client side loads all the data (counts are handled here, after the data is loaded)
                m_RecordCount = data_table.Rows.Count
                m_PageCount = Math.Truncate((m_RecordCount - 1) / m_PageSize) + 1 'leave m_ variables
                loaded_start = If(m_RecordCount = 0, 0, 1)
                loaded_end = m_RecordCount
            Else
                'server side loads one page at a time (counts are handled above)
                If PageSize > 0 Then
                    loaded_start = AbsolutePosition
                    loaded_end = AbsolutePosition + PageSize - 1
                End If
            End If
            If m_RecordCount > 0 AndAlso AbsolutePosition = 0 Then AbsolutePosition = 1
            page_deleted_count = 0

            If m_Fields.Count = 0 Then
                'if this is the first page fetched, then prepare the fields
                dictFields = New Dictionary(Of String, Integer)(data_table.Columns.Count, StringComparer.OrdinalIgnoreCase)

                For i As Integer = 0 To data_table.Columns.Count - 1
                    Dim new_field As Field = New Field() With {.RS = Me, .column = i}
                    m_Fields.Add(new_field, data_table.Columns(i).ColumnName)
                Next
            End If
            m_State = ObjectStateEnum.adStateOpen
        End Sub

        ''' <summary>
        ''' Pull current ADO.NET DataRow object (this function is not part of the original ADODB object)
        ''' </summary>
        ''' <returns>Current ADO.NET DataRow object</returns>
        Public Function GetCurrentRow() As Data.DataRow
            CheckPage()
            Dim row As Data.DataRow = Nothing
            Select Case EditMode
                Case EditModeEnum.adEditAdd
                    row = data_row
                Case Else 'EditModeEnum.adEditInProgress, EditModeEnum.adEditNone, EditModeEnum.adEditDelete
                    row = data_table.Rows(AbsolutePosition - loaded_start)
            End Select
            Return row
        End Function

        Private Sub Recordset_Cleanup()
            If m_ActiveConnection Is Nothing Then Exit Sub 'already been cleaned up

            Try
                m_ActiveConnection.ChildObjects.Remove(Me)
            Catch
            End Try

#If Support_MSSQL Then
            If TypeOf data_adapter Is Data.SqlClient.SqlDataAdapter Then
                Select Case m_CursorLocation
                    Case CursorLocationEnum.adUseServer
                        If cursor_id <> -1 Then
                            Try
                                command.CommandText = "sp_cursorclose"
                                command.CommandType = Data.CommandType.StoredProcedure
                                command.Parameters.Clear()
                                command.Parameters.Add(New Data.SqlClient.SqlParameter("cursor", cursor_id))

                                Using New RunSQLCancellable(command, m_ActiveConnection.CancellationToken)
                                    command.ExecuteNonQuery()
                                End Using
                            Catch
                            End Try
                            cursor_id = -1
                        End If
                End Select
            End If
#End If

#If Support_MSSQL_Compact Then
            If TypeOf data_adapter Is Data.SqlServerCe.SqlCeDataAdapter Then
                Select Case m_CursorLocation
                    Case CursorLocationEnum.adUseServer
                        If cursor_id <> -1 Then
                            Try
                                command.CommandText = "sp_cursorclose"
                                command.CommandType = Data.CommandType.StoredProcedure
                                command.Parameters.Clear()
                                command.Parameters.Add(New Data.SqlServerCe.SqlCeParameter("cursor", cursor_id))

                                Using New RunSQLCancellable(command, m_ActiveConnection.CancellationToken)
                                    command.ExecuteNonQuery()
                                End Using
                            Catch
                            End Try
                            cursor_id = -1
                        End If
                End Select
            End If
#End If

            'dispose of database objects
            Try
                If data_adapter Is Nothing Then
#If Support_MSSQL Then
                ElseIf TypeOf data_adapter Is Data.SqlClient.SqlDataAdapter Then
                    CType(data_adapter, Data.SqlClient.SqlDataAdapter).Dispose()
#End If
#If Support_MSSQL_Compact Then
                ElseIf TypeOf data_adapter Is Data.SqlServerCe.SqlCeDataAdapter Then
                    CType(data_adapter, Data.SqlServerCe.SqlCeDataAdapter).Dispose()
#End If
#If Support_MySQL Then
                ElseIf TypeOf data_adapter Is MySql.Data.MySqlClient.MySqlDataAdapter Then
                    CType(data_adapter, MySql.Data.MySqlClient.MySqlDataAdapter).Dispose()
#End If
#If Support_Oracle Then
                ElseIf TypeOf data_adapter Is Data.OracleClient.OracleDataAdapter Then
                    CType(data_adapter, Data.OracleClient.OracleDataAdapter).Dispose()
#End If
#If Support_SQLite Then
                ElseIf TypeOf data_adapter Is Data.SQLite.SQLiteDataAdapter Then
                    CType(data_adapter, Data.SQLite.SQLiteDataAdapter).Dispose()
#End If
                Else
                    CType(data_adapter, Data.OleDb.OleDbDataAdapter).Dispose()
                End If
                If Not data_table Is Nothing Then data_table.Dispose()
                If Not schema_table Is Nothing Then schema_table.Dispose()
                If Not command Is Nothing Then command.Dispose()
            Catch
                'dispose error
            End Try

            'clean up objects
            data_adapter = Nothing
            data_table = Nothing
            schema_table = Nothing
            schema_table_has_key_info = False
            m_Fields = Nothing
            m_ActiveConnection = Nothing
            data_row = Nothing
            updatedFields = Nothing
            accessedFields = Nothing
            dictFields = Nothing
            command = Nothing
        End Sub

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    If m_State <> ObjectStateEnum.adStateClosed Then Recordset_Cleanup()
                    m_State = ObjectStateEnum.adStateClosed 'shouldn't need this, but just in case someone calls dispose before close
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Public Class Fields
        Implements GenericEnumerable(Of Field)
        Implements IDisposable

        Protected Friend RS As Recordset
        Private m_Field As New Dictionary(Of Integer, Field)()
        Dim position As Integer = -1

        Public Sub New()
        End Sub

        Default Public Property Item(ByVal index As Integer) As Field Implements GenericEnumerable(Of Field).Item
            Get
                RS.CheckPage()
                Return m_Field(index)
            End Get
            Set(ByVal Value As Field)
                RS.CheckPage()
                m_Field(index) = Value
            End Set
        End Property

        Default Public Property Item(ByVal index As String) As Field
            Get
                RS.CheckPage()
                Dim i As Integer
                If RS.dictFields.TryGetValue(index, i) Then
                    Return m_Field(i)
                Else
                    DoError(eErrorType.ERROR_NOT_FOUND, index)
                    Return Nothing
                End If
            End Get
            Set(ByVal Value As Field)
                RS.CheckPage()
                m_Field(RS.dictFields(index)) = Value
            End Set
        End Property

        Public ReadOnly Property Exists(ByVal index As String) As Boolean
            Get
                Return RS.dictFields.ContainsKey(index)
            End Get
        End Property


        Public Sub Add(ByVal value As Field)
            m_Field.Add(m_Field.Count, value)
        End Sub

        Public Sub Add(ByVal value As Field, ByVal name As String)
            If Not RS Is Nothing Then RS.dictFields.Add(name, m_Field.Count)
            m_Field.Add(m_Field.Count, value)
        End Sub

        Public ReadOnly Property Count() As Integer Implements GenericEnumerable(Of Field).Count
            Get
                Return m_Field.Count
            End Get
        End Property

        Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of Field) Implements System.Collections.Generic.IEnumerable(Of Field).GetEnumerator
            Return New GenericEnumerator(Of Field)(Me)
        End Function

        Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return New GenericEnumerator(Of Field)(Me)
        End Function

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Public Class Field
        Implements IDisposable

        Private m_ActualSize As Integer
        Private m_Attributes As Integer
        Private m_DefinedSize As Integer
        Private m_Name As String
        Private m_OriginalValue As Object
        Private m_UnderlyingValue As Object
        Protected Friend RS As Recordset 'the recordset that owns us
        Protected Friend column As Integer 'which field column we are in the row

        Public Sub New()
        End Sub

        Public ReadOnly Property Properties() As Properties
            Get
                Return New Properties() With {.RS = RS, .column = column}
            End Get
        End Property

        Public ReadOnly Property ActualSize() As Integer
            Get
                Return Len(Me.Value) '.ToString.Length
            End Get
        End Property

        Public ReadOnly Property Attributes() As Integer
            Get
                Return m_Attributes
            End Get
        End Property

        Public ReadOnly Property DefinedSize() As Integer
            Get
                Dim ret = RS.data_table.Columns(column).MaxLength
                If ret = -1 Then ret = CType(RS.schema_table.Rows(column).Item("ColumnSize"), Integer)
                Return ret
            End Get
        End Property

        Public ReadOnly Property Name() As String
            Get
                Return RS.data_table.Columns(column).ColumnName
            End Get
        End Property

        Public ReadOnly Property OriginalValue() As Object
            Get
                If RS.State <> ObjectStateEnum.adStateOpen Then DoError(eErrorType.ERROR_CLOSED)
                Dim ret As Object = GetCurrentRow().Item(column, Data.DataRowVersion.Original)
                If TypeOf ret Is Guid Then
                    ret = "{" & ret.ToString() & "}"
                ElseIf TypeOf ret Is Decimal Then
                    If ret = Decimal.Zero Then ret = Decimal.Zero 'we have to do this or else sometimes residual values comes through that can only be seen with Decimal.GetBits(), and you don't get a perfect zero
                End If
                Return ret
            End Get
        End Property

        Public ReadOnly Property UnderlyingValue() As Object
            Get
                If RS.State <> ObjectStateEnum.adStateOpen Then DoError(eErrorType.ERROR_CLOSED)
                Return Value()
            End Get
        End Property

        Property Value() As Object
            Get
                If RS.State <> ObjectStateEnum.adStateOpen Then DoError(eErrorType.ERROR_CLOSED)

                If RS.enableAccessedFieldsWhenReading Then
                    If RS.accessedFields Is Nothing Then RS.accessedFields = New Dictionary(Of Integer, String)()
                    RS.accessedFields(column) = RS.data_table.Columns(column).ColumnName 'add to accessedFields
                End If

                Dim ret As Object = GetCurrentRow().Item(column)
                If TypeOf ret Is Guid Then
                    ret = "{" & ret.ToString() & "}"
                ElseIf TypeOf ret Is Decimal Then
                    If ret = Decimal.Zero Then ret = Decimal.Zero 'we have to do this or else sometimes residual values comes through that can only be seen with Decimal.GetBits(), and you don't get a perfect zero
                End If
                Return ret
            End Get
            Set(ByVal Value As Object)
                If RS.State <> ObjectStateEnum.adStateOpen Then DoError(eErrorType.ERROR_CLOSED)
                If RS.LockType = LockTypeEnum.adLockReadOnly Then DoError(eErrorType.ERROR_NOUPDATES)

                Dim row As Data.DataRow = Nothing
                Select Case RS.EditMode
                    Case EditModeEnum.adEditNone
                        RS.updatedFields = New Dictionary(Of Integer, String)()
                        RS.accessedFields = New Dictionary(Of Integer, String)()
                        RS.m_EditMode = EditModeEnum.adEditInProgress
                        row = GetCurrentRow()
                        row.BeginEdit()
                    Case Else
                        row = GetCurrentRow()
                End Select
                Try
                    Dim oldValue, newValue As Object
                    oldValue = row.Item(column)

                    Dim typ As Type = row.Table.Columns(column).DataType
                    If typ.Equals(GetType(Guid)) AndAlso TypeOf Value Is String Then
                        'VB6 uses string for GUIDS, so we have to convert
                        newValue = New Guid(Value.ToString())
                    ElseIf typ.Equals(GetType(Date)) AndAlso TypeOf Value Is Double Then
                        'VB6 allows doubles and dates to flip back and forth, where .NET requires a specific .FromOADate() conversion
                        newValue = Date.FromOADate(Value)
                    Else
                        newValue = Value
                    End If
                    row.Item(column) = newValue

                    'update the updatedFields Dictionary with the value that we just modified
                    If Not RS.updatedFields Is Nothing Then
                        RS.accessedFields.Item(column) = RS.data_table.Columns(column).ColumnName 'always add to accessedFields

                        'add the field to the updatedFields collection,
                        'unless we are EDITING a record and the new field value is the same as the old field value, then skip (optimization)
                        If RS.EditMode <> EditModeEnum.adEditInProgress OrElse oldValue?.ToString() <> newValue?.ToString() Then
                            RS.updatedFields(column) = RS.data_table.Columns(column).ColumnName
                        End If
                    End If
                Catch
                End Try
            End Set
        End Property

        Public ReadOnly Property Type() As DataTypeEnum
            Get
                'If RS.schema_table Is Nothing Then Return DataTypeEnum.adIUnknown
                If False Then
                    Return Nothing
#If Support_MSSQL Then
                ElseIf TypeOf RS.ActiveConnection.connection Is Data.SqlClient.SqlConnection Then
                    Return TranslateType(CType(RS.schema_table.Rows(column).Item("ProviderType"), Data.SqlDbType), CStr(RS.schema_table.Rows(column).Item("DataTypeName")))
#End If
#If Support_MSSQL_Compact Then
                ElseIf TypeOf RS.ActiveConnection.connection Is Data.SqlServerCe.SqlCeConnection Then
                    Return TranslateType(CType(RS.schema_table.Rows(column).Item("ProviderType"), Data.SqlDbType), CStr(RS.schema_table.Rows(column).Item("DataTypeName")))
#End If
                Else
                    Return CType(RS.schema_table.Rows(column).Item("ProviderType"), DataTypeEnum)
                End If
            End Get
        End Property

        Public ReadOnly Property UdtTypeName() As String
            Get
                Try
                    Dim nm As String = RS.schema_table.Rows(column).Item("ProviderSpecificDataType").Name
                    Select Case nm
                        Case "SqlGeography"
                            Return "Geography"
                        Case Else
                            Return nm
                    End Select
                Catch ex As Exception
                End Try

                Return Nothing
            End Get
        End Property

        Private Function GetCurrentRow() As Data.DataRow
            Dim row As Data.DataRow = Nothing
            Select Case RS.EditMode
                Case EditModeEnum.adEditAdd
                    row = RS.data_row
                Case Else 'EditModeEnum.adEditInProgress, EditModeEnum.adEditNone, EditModeEnum.adEditDelete
                    row = RS.data_table.Rows(RS.AbsolutePosition - RS.loaded_start)
            End Select
            Return row
        End Function

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    'Refer to SqlDataReader.GetSchemaTable Method for information on each schema enumeration type
    'http://msdn.microsoft.com/en-us/library/system.data.sqlclient.sqldatareader.getschematable(v=vs.110).aspx
    Public Enum eRECORDSET_PROPERTIES
        AllowDBNull = 1
        BaseCatalogName = 2
        BaseColumnName = 3
        BaseSchemaName = 4
        BaseServerName = 5
        BaseTableName = 6
        ColumnName = 7
        ColumnOrdinal = 8
        ColumnSize = 9
        DataTypeName = 10
        IsAliased = 11
        IsAutoIncrement = 12
        IsColumnSet = 13
        IsExpression = 14
        IsHidden = 15
        IsIdentity = 16
        IsKey = 17
        IsLong = 18
        IsReadOnly = 19
        IsRowVersion = 20
        IsUnique = 21
        NonVersionedProviderType = 22
        NumericPrecision = 23
        NumericScale = 24
        ProviderSpecificDataType = 25
        ProviderType = 26
        UdtAssemblyQualifiedName = 27
        XmlSchemaCollectionDatabase = 28
        XmlSchemaCollectionName = 29
        XmlSchemaCollectionOwningSchema = 30
    End Enum

    Public Class Properties
        Implements GenericEnumerableReadOnly(Of [Property])
        Implements IDisposable

        Private position As Integer = -1
        Protected Friend RS As Recordset
        Protected Friend column As Integer = 0 'tells which column we are looking for information about (which equaltes to a row in the schema table)

        Public Sub New()
        End Sub

        Public ReadOnly Property Count() As Integer Implements GenericEnumerableReadOnly(Of [Property]).Count
            Get
                Return RS.schema_table.Columns.Count
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal index As Integer) As [Property] Implements GenericEnumerableReadOnly(Of [Property]).Item
            Get
                Return New [Property] With {.m_Name = RS.schema_table.Columns(index).ColumnName, .m_Value = RS.schema_table.Rows(column).Item(index)}
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal index As String) As [Property]
            Get
                Dim pull_schema_with_keys As Boolean = False
                Select Case UCase(index) 'some properties have been renamed from ADODB to ADO.NET
                    Case "KEYCOLUMN"
                        index = "IsKey"
                        If Not RS.schema_table_has_key_info Then pull_schema_with_keys = True
                    Case "ISKEY", "ISUNIQUE"
                        If Not RS.schema_table_has_key_info Then pull_schema_with_keys = True
                End Select
                'if we need information about the primary keys, we need to repull the schema with the extra key information
                If pull_schema_with_keys Then
                    RS.schema_table_has_key_info = True
                    Try
                        If Not RS.schema_table Is Nothing Then
                            RS.schema_table.Dispose()
                            RS.schema_table = Nothing
                        End If
                        Using rdr = RS.command.ExecuteReader(Data.CommandBehavior.KeyInfo Or Data.CommandBehavior.SchemaOnly)
                            RS.schema_table = rdr.GetSchemaTable()
                            rdr.Close()
                        End Using
                        'make sure the data_table.TableName if filled in... if not, pull it from the first column's basetablename property
                        If Len(RS.data_table.TableName) = 0 Then RS.data_table.TableName = RS.schema_table.Rows(0).Item("BaseTableName")
                    Catch
                        Return New [Property]()
                    End Try
                End If
                Return New [Property] With {.m_Name = RS.schema_table.Columns(index).ColumnName, .m_Value = RS.schema_table.Rows(column).Item(index)}
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal index As eRECORDSET_PROPERTIES) As [Property]
            Get
                Return Item(index.ToString())
            End Get
        End Property

        Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of [Property]) Implements System.Collections.Generic.IEnumerable(Of [Property]).GetEnumerator
            Return New GenericEnumerator(Of [Property])(Me)
        End Function

        Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return New GenericEnumerator(Of [Property])(Me)
        End Function

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Public Class [Property]
        Implements IDisposable

        Protected Friend m_Name As String
        Protected Friend m_Value As Object

        Public ReadOnly Property Attributes() As Integer
            Get
                Return 1
            End Get
        End Property

        Public ReadOnly Property Name() As String
            Get
                Return UCase(m_Name)
            End Get
        End Property

        Public ReadOnly Property Value() As Object
            Get
                If m_Value Is DBNull.Value Then
                    Return ""
                Else
                    Return m_Value
                End If
            End Get
        End Property

        Public ReadOnly Property Type() As DataTypeEnum
            Get
                Return TranslateType(m_Value.GetType())
            End Get
        End Property

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Public Class Command
        Implements IDisposable

        Private m_Command As Data.IDbCommand
        Private m_CommandText As String
        Private m_CommandTimeout As Integer
        Private m_CommandType As CommandTypeEnum

        Private m_ActiveConnection As Connection
        Private m_Parameters As New Parameters()
        Private m_Prepared As Boolean

        Public Sub New()
        End Sub

        Public Property ActiveConnection() As Connection
            Get
                Return m_ActiveConnection
            End Get
            Set(ByVal Value As Connection)
                m_ActiveConnection = Value
                m_Command = Value.Get_ProviderSpecific_Command()
                m_Command.Connection = Value.connection
            End Set
        End Property

        Public Property CommandText() As String
            Get
                Return m_CommandText
            End Get
            Set(ByVal Value As String)
                m_CommandText = Value
            End Set
        End Property

        Public Property CommandStream() As String 'not supported
            Get
                Throw New NotImplementedException()
            End Get
            Set(ByVal Value As String)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Property CommandTimeout() As Integer
            Get
                Return m_CommandTimeout
            End Get
            Set(ByVal Value As Integer)
                m_CommandTimeout = Value
            End Set
        End Property

        Public Property CommandType() As CommandTypeEnum
            Get
                Select Case m_CommandType
                    Case Data.CommandType.StoredProcedure
                        Return CommandTypeEnum.adCmdStoredProc
                    Case Data.CommandType.TableDirect
                        Return CommandTypeEnum.adCmdTable
                    Case Else 'Data.CommandType.Text
                        Return CommandTypeEnum.adCmdText
                End Select
            End Get
            Set(ByVal Value As CommandTypeEnum)
                Select Case Value
                    Case CommandTypeEnum.adCmdText
                        m_CommandType = Data.CommandType.Text
                    Case CommandTypeEnum.adCmdTable, CommandTypeEnum.adCmdTableDirect
                        m_CommandType = Data.CommandType.TableDirect
                    Case CommandTypeEnum.adCmdStoredProc
                        m_CommandType = Data.CommandType.StoredProcedure
                    Case Else 'CommandTypeEnum.adCmdText, CommandTypeEnum.adCmdFile, CommandTypeEnum.adCmdUnknown, CommandTypeEnum.adCmdUnspecified
                        m_CommandType = Data.CommandType.Text
                End Select
            End Set
        End Property

        Public Property Prepared() As Boolean
            Get
                Return m_Prepared
            End Get
            Set(ByVal value As Boolean)
                m_Prepared = True
            End Set
        End Property

        Public Function CreateParameter(Optional ByVal Name As String = "", Optional ByVal Type As DataTypeEnum = DataTypeEnum.adEmpty, Optional ByVal Direction As ParameterDirectionEnum = ParameterDirectionEnum.adParamInput, Optional ByVal Size As Integer = 0, Optional ByVal Value As Object = Nothing) As Parameter
            Return New Parameter() With {.Name = Name, .Type = Type, .Direction = Direction, .Size = Size, .Value = Value}
        End Function

        Public ReadOnly Property Parameters() As Parameters
            Get
                Return m_Parameters
            End Get
        End Property

        'TODO: add support to the Options parameter for ExecuteOptionEnum
        Function Execute(Optional ByRef RecordsAffected As Integer = 0, Optional ByVal Parameters() As Parameter = Nothing, Optional ByVal Options As CommandTypeEnum = CommandTypeEnum.adCmdUnspecified) As Recordset
            If m_Command Is Nothing OrElse m_ActiveConnection Is Nothing OrElse m_ActiveConnection.State = ObjectStateEnum.adStateClosed Then DoError(eErrorType.ERROR_CONNECTION_CLOSED)
            m_Command.CommandTimeout = m_CommandTimeout
            m_Command.CommandType = m_CommandType
            m_Command.CommandText = m_CommandText

            Dim RS As New Recordset()
            RS.ActiveConnection = m_ActiveConnection
            RS.InitOpen(m_Command, m_ActiveConnection, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly) 'as far as I can tell, Execute() only uses one type of cursor/lock
            m_Command.Parameters.Clear()
            If Not Parameters Is Nothing Then
                For Each p In Parameters
                    m_Command.Parameters.Add(p.m_Parameter)
                Next
            Else
                For Each p In m_Parameters
                    m_Command.Parameters.Add(p.m_Parameter)
                Next
            End If
            If Options <> CommandTypeEnum.adCmdUnspecified Then CommandType = Options

            'Only OleDBConnection supports positional parameters using question marks (i.e. where x=?), but provider specific only support named parameters (where x=@pname), so we have to go through and replace the ? with @pname
            If Not TypeOf m_Command Is Data.OleDb.OleDbConnection Then
                If m_Command.Parameters.Count > 0 Then
                    Dim tmp_commandtext As String = m_Command.CommandText
                    Dim pname As String
                    Dim pos As Integer = 1
                    Dim pcount As Integer = 0
                    Do
                        pos = InStr(pos, tmp_commandtext, "?")
                        If pos = 0 Then Exit Do
                        pname = m_Command.Parameters(pcount).ParameterName
                        tmp_commandtext = Replace(tmp_commandtext, "?", "@" & pname, Count:=1)
                        pos += Len(pname) + 1
                        pcount += 1
                    Loop
                    m_Command.CommandText = tmp_commandtext
                End If
            End If
            If m_Prepared Then m_Command.Prepare()
            RS.LoadPage(m_Command)
            RecordsAffected = RS.RecordsAffected
            Return RS
        End Function

        Public Property Name() As String
            Get
                Throw New NotImplementedException()
            End Get
            Set(ByVal Value As String)
                Throw New NotImplementedException()
            End Set
        End Property

        Public ReadOnly Property State() As ObjectStateEnum
            Get
                If Not m_ActiveConnection Is Nothing Then
                    Return m_ActiveConnection.State
                Else
                    Return ObjectStateEnum.adStateClosed
                End If
            End Get
        End Property

        'TODO: add support to the Options parameter for ExecuteOptionEnum
        Public Sub Cancel()
        End Sub

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    Try
                        If Not m_Command Is Nothing Then m_Command.Dispose()
                    Catch
                    Finally
                        m_Command = Nothing
                    End Try
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Public Class Parameters
        Implements GenericEnumerableReadOnly(Of Parameter)
        Implements IDisposable

        Protected Friend m_Parameters As New Collections.Generic.List(Of Parameter)
        Private position As Integer = -1

        Public Sub Append(ByVal [Object] As Parameter)
            m_Parameters.Add([Object])
        End Sub

        Public Sub Delete(ByVal Index As Integer)
            m_Parameters.RemoveAt(Index)
        End Sub

        Public ReadOnly Property Count() As Integer Implements GenericEnumerableReadOnly(Of Parameter).Count
            Get
                Return m_Parameters.Count
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal index As Integer) As Parameter Implements GenericEnumerableReadOnly(Of Parameter).Item
            Get
                Return m_Parameters(index)
            End Get
        End Property

        Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of Parameter) Implements System.Collections.Generic.IEnumerable(Of Parameter).GetEnumerator
            Return New GenericEnumeratorReadOnly(Of Parameter)(Me)
        End Function

        Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return New GenericEnumeratorReadOnly(Of Parameter)(Me)
        End Function

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Public Class Parameter
        Implements IDisposable

        Protected Friend m_Parameter As New Data.SqlClient.SqlParameter()

        Public Property Name() As String
            Get
                Return m_Parameter.ParameterName
            End Get
            Set(ByVal value As String)
                m_Parameter.ParameterName = value
            End Set
        End Property

        Public Property Direction() As ParameterDirectionEnum
            Get
                Select Case m_Parameter.Direction
                    Case Data.ParameterDirection.Input
                        Return ParameterDirectionEnum.adParamInput
                    Case Data.ParameterDirection.Output
                        Return ParameterDirectionEnum.adParamOutput
                    Case Data.ParameterDirection.InputOutput
                        Return ParameterDirectionEnum.adParamInputOutput
                    Case Data.ParameterDirection.ReturnValue
                        Return ParameterDirectionEnum.adParamReturnValue
                    Case Else
                        Return ParameterDirectionEnum.adParamUnknown
                End Select
            End Get
            Set(ByVal value As ParameterDirectionEnum)
                Select Case value
                    Case ParameterDirectionEnum.adParamInput
                        m_Parameter.Direction = Data.ParameterDirection.Input
                    Case ParameterDirectionEnum.adParamOutput
                        m_Parameter.Direction = Data.ParameterDirection.Output
                    Case ParameterDirectionEnum.adParamInputOutput
                        m_Parameter.Direction = Data.ParameterDirection.InputOutput
                    Case ParameterDirectionEnum.adParamReturnValue
                        m_Parameter.Direction = Data.ParameterDirection.ReturnValue
                    Case ParameterDirectionEnum.adParamUnknown
                        m_Parameter.Direction = 0
                End Select
            End Set
        End Property

        Public Property NumericScale() As Byte
            Get
                Return m_Parameter.Scale
            End Get
            Set(ByVal value As Byte)
                m_Parameter.Scale = value
            End Set
        End Property

        Public Property Precision() As Byte
            Get
                Return m_Parameter.Precision
            End Get
            Set(ByVal value As Byte)
                m_Parameter.Precision = value
            End Set
        End Property

        Public Property Size() As Integer
            Get
                Return m_Parameter.Size
            End Get
            Set(ByVal value As Integer)
                m_Parameter.Size = value
            End Set
        End Property

        Public Property Type() As DataTypeEnum
            Get
                Return TranslateType(m_Parameter.SqlDbType)
            End Get
            Set(ByVal value As DataTypeEnum)
                m_Parameter.SqlDbType = TranslateType(value)
            End Set
        End Property

        Public Property Value() As Object
            Get
                Return m_Parameter.Value
            End Get
            Set(ByVal value As Object)
                m_Parameter.Value = value
            End Set
        End Property

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Private Sub DoError(ByVal num As eErrorType, Optional extraInfo As String = Nothing)
        Select Case num
            Case eErrorType.ERROR_OUT_OF_BOUNDS
                Err.Raise(num, , "Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record.")
            Case eErrorType.ERROR_CONTEXT
                Err.Raise(num, , "Operation is not allowed in this context.")
            Case eErrorType.ERROR_CLOSED
                Err.Raise(num, , "Operation is not allowed when the object is closed.")
            Case eErrorType.ERROR_OPEN
                Err.Raise(num, , "Operation is not allowed when the object is open.")
            Case eErrorType.ERROR_NOT_FOUND
                Err.Raise(num, , "Item cannot be found in the collection corresponding to the requested name or ordinal." & If(Len(extraInfo) > 0, ".. Field=" & extraInfo, ""))
            Case eErrorType.ERROR_NOUPDATES
                Err.Raise(num, , "Current Recordset does not support updating.  This may be a limitation of the provider, or of the selected locktype.")
            Case eErrorType.ERROR_CONNECTION_CLOSED
                Err.Raise(num, , "The connection cannot be used to perform this operation. It is either closed or invalid in this context.")
            Case eErrorType.ERROR_NO_TRANS
                Err.Raise(num, , "No transaction is active.")
        End Select
    End Sub

    'translates a .NET type to the DataTypeEnum
    Private Function TranslateType(ByVal columnType As Type) As DataTypeEnum
        Select Case columnType.UnderlyingSystemType.ToString()
            Case "System.Guid"
                Return ADODB.DataTypeEnum.adGUID

            Case "System.Boolean"
                Return ADODB.DataTypeEnum.adBoolean

            Case "System.Byte"
                Return ADODB.DataTypeEnum.adUnsignedTinyInt

            Case "System.Char"
                Return ADODB.DataTypeEnum.adChar

            Case "System.DateTime"
                Return ADODB.DataTypeEnum.adDate

            Case "System.Decimal"
                Return ADODB.DataTypeEnum.adDecimal

            Case "System.Double"
                Return ADODB.DataTypeEnum.adDouble

            Case "System.Int16"
                Return ADODB.DataTypeEnum.adSmallInt

            Case "System.Int32"
                Return ADODB.DataTypeEnum.adInteger

            Case "System.Int64"
                Return ADODB.DataTypeEnum.adBigInt

            Case "System.SByte"
                Return ADODB.DataTypeEnum.adTinyInt

            Case "System.Single"
                Return ADODB.DataTypeEnum.adSingle

            Case "System.UInt16"
                Return ADODB.DataTypeEnum.adUnsignedSmallInt

            Case "System.UInt32"
                Return ADODB.DataTypeEnum.adUnsignedInt

            Case "System.UInt64"
                Return ADODB.DataTypeEnum.adUnsignedBigInt

            Case Else '"System.String"
                Return ADODB.DataTypeEnum.adVarChar
        End Select
    End Function

    'translates a SQL data type string to the DataTypeEnum
    Public Function TranslateType(ByVal dataType As Data.SqlDbType, Optional ByVal dataTypeName As String = "") As DataTypeEnum
        Select Case dataType
            Case Data.SqlDbType.UniqueIdentifier
                Return DataTypeEnum.adGUID

            Case Data.SqlDbType.Bit
                Return DataTypeEnum.adBoolean

            Case Data.SqlDbType.TinyInt
                If LCase(Left(dataTypeName, 1)) = "u" Then
                    Return DataTypeEnum.adUnsignedTinyInt
                Else
                    Return DataTypeEnum.adTinyInt
                End If

            Case Data.SqlDbType.SmallInt
                If LCase(Left(dataTypeName, 1)) = "u" Then
                    Return DataTypeEnum.adUnsignedSmallInt
                Else
                    Return DataTypeEnum.adSmallInt
                End If

            Case Data.SqlDbType.Int
                If LCase(Left(dataTypeName, 1)) = "u" Then
                    Return DataTypeEnum.adUnsignedInt
                Else
                    Return DataTypeEnum.adInteger
                End If

            Case Data.SqlDbType.BigInt
                If LCase(Left(dataTypeName, 1)) = "u" Then
                    Return DataTypeEnum.adUnsignedBigInt
                Else
                    Return DataTypeEnum.adBigInt
                End If

            Case Data.SqlDbType.Money, Data.SqlDbType.SmallMoney
                Return DataTypeEnum.adCurrency

            Case Data.SqlDbType.Real
                Return DataTypeEnum.adSingle

            Case Data.SqlDbType.Float
                Return DataTypeEnum.adDouble

            Case Data.SqlDbType.Decimal
                Return DataTypeEnum.adDecimal

            Case Data.SqlDbType.SmallDateTime, Data.SqlDbType.Date, Data.SqlDbType.DateTime, Data.SqlDbType.DateTime2, Data.SqlDbType.DateTimeOffset, Data.SqlDbType.Time
                Return DataTypeEnum.adDBTimeStamp

            Case Data.SqlDbType.DateTimeOffset
                Return DataTypeEnum.adDBTimeStampOffset

            Case Data.SqlDbType.Char
                Return DataTypeEnum.adChar

            Case Data.SqlDbType.NChar
                Return DataTypeEnum.adWChar

            Case Data.SqlDbType.VarChar
                Return DataTypeEnum.adVarChar

            Case Data.SqlDbType.NVarChar
                Return DataTypeEnum.adVarWChar

            Case Data.SqlDbType.Text
                Return DataTypeEnum.adLongVarChar

            Case Data.SqlDbType.NText
                Return DataTypeEnum.adLongVarWChar

            Case Data.SqlDbType.Binary
                Return DataTypeEnum.adVarBinary

            Case Data.SqlDbType.VarBinary
                Return DataTypeEnum.adVarBinary

            Case Data.SqlDbType.Image
                Return DataTypeEnum.adBinary

            Case Data.SqlDbType.Timestamp
                Return DataTypeEnum.adDBTimeStamp

            Case Data.SqlDbType.Variant
                Return DataTypeEnum.adVariant

            Case Data.SqlDbType.Xml
                Return DataTypeEnum.adVarWChar

            Case Data.SqlDbType.Udt
                Return DataTypeEnum.adUserDefined

            Case Else
                Throw New Exception("Unknown type")
        End Select
    End Function

    'translates a SQL data type string to the DataTypeEnum
    Private Function TranslateType(ByVal dataType As DataTypeEnum) As Data.SqlDbType
        Select Case dataType
            Case DataTypeEnum.adGUID
                Return Data.SqlDbType.UniqueIdentifier

            Case DataTypeEnum.adBoolean
                Return Data.SqlDbType.Bit

            Case DataTypeEnum.adUnsignedTinyInt, DataTypeEnum.adTinyInt
                Return Data.SqlDbType.TinyInt

            Case DataTypeEnum.adUnsignedSmallInt, DataTypeEnum.adSmallInt
                Return Data.SqlDbType.SmallInt

            Case DataTypeEnum.adUnsignedInt, DataTypeEnum.adInteger
                Return Data.SqlDbType.Int

            Case DataTypeEnum.adUnsignedBigInt, DataTypeEnum.adBigInt
                Return Data.SqlDbType.BigInt

            Case DataTypeEnum.adCurrency
                Return Data.SqlDbType.Money

            Case DataTypeEnum.adSingle
                Return Data.SqlDbType.Real

            Case DataTypeEnum.adDouble
                Return Data.SqlDbType.Float

            Case DataTypeEnum.adDecimal
                Return Data.SqlDbType.Decimal

            Case DataTypeEnum.adNumeric
                Return Data.SqlDbType.Decimal

            Case DataTypeEnum.adChar
                Return Data.SqlDbType.Char

            Case DataTypeEnum.adWChar
                Return Data.SqlDbType.NChar

            Case DataTypeEnum.adVarChar
                Return Data.SqlDbType.VarChar

            Case DataTypeEnum.adVarWChar, DataTypeEnum.adBSTR
                Return Data.SqlDbType.NVarChar

            Case DataTypeEnum.adLongVarChar
                Return Data.SqlDbType.Text

            Case DataTypeEnum.adLongVarWChar
                Return Data.SqlDbType.NText

            Case DataTypeEnum.adLongVarBinary, DataTypeEnum.adVarBinary, DataTypeEnum.adBinary
                Return Data.SqlDbType.VarBinary

            Case DataTypeEnum.adDBDate, DataTypeEnum.adDate
                Return Data.SqlDbType.Date

            Case DataTypeEnum.adDBTime, DataTypeEnum.adDBTimeStamp, DataTypeEnum.adDBFileTime, DataTypeEnum.adFileTime
                Return Data.SqlDbType.DateTime

            Case DataTypeEnum.adDBTimeStampOffset
                Return Data.SqlDbType.DateTimeOffset

            Case DataTypeEnum.adVariant
                Return Data.SqlDbType.Variant

            Case DataTypeEnum.adEmpty 'no mapping, so using NVarChar for now
                Return Data.SqlDbType.NVarChar

            Case DataTypeEnum.adUserDefined
                Return Data.SqlDbType.Udt

            Case Else
                Throw New Exception("Unknown type")
        End Select
    End Function

    'GENERIC ENUMERABLE CLASSES
    Private Interface GenericEnumerable(Of T)
        Inherits Collections.Generic.IEnumerable(Of T)

        Property Item(ByVal index As Integer) As T
        ReadOnly Property Count() As Integer
    End Interface

    Private Interface GenericEnumerableReadOnly(Of T)
        Inherits Collections.Generic.IEnumerable(Of T)

        ReadOnly Property Item(ByVal index As Integer) As T
        ReadOnly Property Count() As Integer
    End Interface

    Private Class GenericEnumeratorReadOnly(Of T)
        Implements Collections.Generic.IEnumerator(Of T)
        Implements IDisposable

        Private position As Integer = -1
        Private parent As GenericEnumerableReadOnly(Of T)

        Public Sub New(ByVal Enumerable As GenericEnumerableReadOnly(Of T))
            parent = Enumerable
        End Sub

        Public ReadOnly Property Current() As T Implements System.Collections.Generic.IEnumerator(Of T).Current
            Get
                Return parent.Item(position)
            End Get
        End Property

        Public ReadOnly Property Current1() As Object Implements System.Collections.IEnumerator.Current
            Get
                Return parent.Item(position)
            End Get
        End Property

        Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
            position += 1
            Return (position < parent.Count)
        End Function

        Public Sub Reset() Implements System.Collections.IEnumerator.Reset
            position = -1
        End Sub

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Private Class GenericEnumerator(Of T)
        Implements Collections.Generic.IEnumerator(Of T)
        Implements IDisposable

        Private position As Integer = -1
        Private parent As GenericEnumerable(Of T)

        Public Sub New(ByVal Enumerable As GenericEnumerable(Of T))
            parent = Enumerable
        End Sub

        Public ReadOnly Property Current() As T Implements System.Collections.Generic.IEnumerator(Of T).Current
            Get
                Return parent.Item(position)
            End Get
        End Property

        Public ReadOnly Property Current1() As Object Implements System.Collections.IEnumerator.Current
            Get
                Return parent.Item(position)
            End Get
        End Property

        Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
            position += 1
            Return (position < parent.Count)
        End Function

        Public Sub Reset() Implements System.Collections.IEnumerator.Reset
            position = -1
        End Sub

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class
End Module
