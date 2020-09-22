Imports Npgsql
Imports System.IO
Public Class DbAdapter
    Implements IDisposable
    Dim mytransaction As NpgsqlTransaction

    Dim _ConnectionStringDict As Dictionary(Of String, String)
    Dim _connectionstring As String
    Private CopyIn1 As NpgsqlCopyIn
    Dim _userid As String
    Dim _password As String

    Public Shared myInstance As DbAdapter

    Public Shared Function getInstance() As DbAdapter
        If myInstance Is Nothing Then
            myInstance = New DbAdapter
        End If
        Return myInstance
    End Function

    Public ReadOnly Property userid As String
        Get
            Return _userid
        End Get
    End Property
    Public ReadOnly Property password As String
        Get
            Return _password
        End Get
    End Property

    Public Property Connectionstring As String
        Get
            Return _connectionstring

        End Get
        Set(ByVal value As String)
            _connectionstring = value
        End Set
    End Property

    Public Sub New()
        InitConnectionStringDict()
        _connectionstring = getConnectionString()
    End Sub

    Public ReadOnly Property ConnectionStringDict As Dictionary(Of String, String)
        Get
            Return _ConnectionStringDict
        End Get
    End Property

    Private Sub InitConnectionStringDict()
        _ConnectionStringDict = New Dictionary(Of String, String)
        Dim connectionstring = getConnectionString()
        Dim connectionstrings() As String = connectionstring.Split(";")
        For i = 0 To (connectionstrings.Length - 1)
            Dim mystrs() As String = connectionstrings(i).Split("=")
            _ConnectionStringDict.Add(mystrs(0), mystrs(1))
        Next i

    End Sub
    Public Function isSalesTeam(ByVal userid) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr = String.Format("select userid from sales.tb_user where lower(userid) = '{0}'", userid.ToString.ToLower)
        If ExScallar(sqlstr) <> "" Then
            myret = True

        End If
        isSalesTeam = myret
        Return myret
    End Function
    Private Function getConnectionString() As String
        _userid = "admin"
        _password = "admin"
        Dim builder As New NpgsqlConnectionStringBuilder()
        builder.ConnectionString = My.Settings.Connectionstring1
        builder.Add("User Id", _userid)
        builder.Add("password", _password)
        builder.Add("CommandTimeout", "1000")
        builder.Add("TimeOut", "1000")
        Return builder.ConnectionString
    End Function

    Public Function ExNonQuery(ByVal sqlstr As String) As Long
        Dim myRet As Long
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                myRet = command.ExecuteNonQuery
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExScallar(ByVal sqlstr As String) As Object
        Dim myRet As Object
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                myRet = command.ExecuteScalar
            End Using
        End Using
        Return myRet
    End Function

    Public Function getDataSet(ByVal sqlstr As String, ByRef DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                'DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                DataAdapter.Fill(DataSet)
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
    Public Function copy(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        result = False
        Dim myReturn As String = ""
        'Convert string to MemoryStream
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.ASCII.GetBytes(InputString))
        Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(InputString))
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.Default.GetBytes(InputString))
        Dim buf(9) As Byte
        Dim CopyInStream As Stream = Nothing
        Dim i As Long
        Using conn = New NpgsqlConnection(getConnectionString())
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                CopyIn1 = New NpgsqlCopyIn(command, conn)
                Try
                    CopyIn1.Start()
                    CopyInStream = CopyIn1.CopyStream
                    i = MemoryStream1.Read(buf, 0, buf.Length)
                    While i > 0
                        CopyInStream.Write(buf, 0, i)
                        i = MemoryStream1.Read(buf, 0, buf.Length)
                        Application.DoEvents()
                    End While
                    CopyInStream.Close()
                    result = True
                Catch ex As NpgsqlException
                    Try
                        CopyIn1.Cancel("Undo Copy")
                        myReturn = ex.Message
                    Catch ex2 As NpgsqlException
                        If ex2.Message.Contains("Undo Copy") Then
                            myReturn = ex2.Message
                        End If
                    End Try
                End Try

            End Using
        End Using

        Return myReturn
    End Function
    Public Function copyUTF(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        result = False
        Dim myReturn As String = ""
        'Convert string to MemoryStream
        Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(InputString))
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.Default.GetBytes(InputString))
        Dim buf(9) As Byte
        Dim CopyInStream As Stream = Nothing
        Dim i As Long
        Using conn = New NpgsqlConnection(getConnectionString())
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                CopyIn1 = New NpgsqlCopyIn(command, conn)
                Try
                    CopyIn1.Start()
                    CopyInStream = CopyIn1.CopyStream
                    i = MemoryStream1.Read(buf, 0, buf.Length)
                    While i > 0
                        CopyInStream.Write(buf, 0, i)
                        i = MemoryStream1.Read(buf, 0, buf.Length)
                        Application.DoEvents()
                    End While
                    CopyInStream.Close()
                    result = True
                Catch ex As NpgsqlException
                    Try
                        CopyIn1.Cancel("Undo Copy")
                        myReturn = ex.Message
                    Catch ex2 As NpgsqlException
                        If ex2.Message.Contains("Undo Copy") Then
                            myReturn = ex2.Message
                        End If
                    End Try
                End Try

            End Using
        End Using

        Return myReturn
    End Function


    Function loglogin(ByVal applicationname As String, ByVal userid As String, ByVal username As String, ByVal computername As String, ByVal time_stamp As Date)
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertlogonhistory", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = applicationname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = computername
            result = cmd.ExecuteNonQuery
        End Using
        Return result
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

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

    Public Sub onRowInsertUpdate(ByVal sender As Object, ByVal e As NpgsqlRowUpdatedEventArgs)
        'Table with autoincrement
        If e.StatementType = StatementType.Insert Or e.StatementType = StatementType.Update Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub
    Public Sub onRowUpdate(ByVal sender As Object, ByVal e As NpgsqlRowUpdatedEventArgs)
        'Table without autoincrement
        If e.StatementType = StatementType.Update Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub
    Function ProductFamilyTX(ByVal familyDataAdapter As FamilyDataAdapter, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)

        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            sqlstr = "sales.sp_updateproductfamily"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "id").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyname").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "type").SourceVersion = DataRowVersion.Current            
            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_insertproductfamily"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyname").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "type").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "id").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_deleteproductfamily"
            DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            DataAdapter.InsertCommand.Transaction = mytransaction
            DataAdapter.UpdateCommand.Transaction = mytransaction
            DataAdapter.DeleteCommand.Transaction = mytransaction

            mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Function MLATX(ByVal mLAAdapter As MLAAdapter, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)

        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            sqlstr = "sales.sp_updatemla"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "id").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "mlaname").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "countryid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "countryname").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "distchannel").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "distchanneldesc").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "mlatype").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validfrom").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validto").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "status").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_insertmla"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "mlaname").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "countryid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "countryname").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "distchannel").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "distchanneldesc").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "mlatype").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validfrom").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validto").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "status").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "id").SourceVersion = DataRowVersion.Current            
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_deletemla"
            DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            DataAdapter.InsertCommand.Transaction = mytransaction
            DataAdapter.UpdateCommand.Transaction = mytransaction
            DataAdapter.DeleteCommand.Transaction = mytransaction

            mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function
    Public Function getConnection() As NpgsqlConnection
        If IsNothing(_userid) Or IsNothing(_password) Then
            Throw New DbAdapterExeption("User Id or Password is blank.")
        End If
        Return New NpgsqlConnection(_connectionstring)
    End Function

    Public Function getDbDataAdapter() As NpgsqlDataAdapter
        Return New NpgsqlDataAdapter
    End Function

    Public Function getCommandObject() As NpgsqlCommand
        Return New NpgsqlCommand
    End Function


    Public Function getCommandObject(ByVal sqlstr As String, ByVal connection As Object) As NpgsqlCommand
        Return New NpgsqlCommand(sqlstr, connection)
    End Function

    Public Function ExecuteScalar(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteScalar
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function


End Class
Public Class DbAdapterExeption
    Inherits ApplicationException
    Public Sub New(ByVal errormessage As String)
        MyBase.New(errormessage)
    End Sub
End Class