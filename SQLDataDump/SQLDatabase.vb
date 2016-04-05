Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Аргумент события смена таблицы для экспорта
''' </summary>
Public Class TableExportEventArgs
    Inherits EventArgs

    ''' <summary>
    ''' Имя экспортируемой таблицы
    ''' </summary>
    Public Property TableName As String

    Public Sub New(name As String)
        Me.TableName = name
    End Sub

End Class

''' <summary>
''' Представляет базу данных SQL Server
''' </summary>
Public Class SQLDatabase
    Implements IDisposable

    Private Const SchemaSql As String = _
        "SELECT T.Name, C.Name, TY.Name FROM sys.tables AS T " & _
            "INNER JOIN sys.columns AS C ON T.object_id = C.object_id " & _
            "INNER JOIN sys.types AS TY " & _
                "ON C.system_type_id = TY.system_type_id " & _
            "WHERE T.type = 'U' " & _
            "ORDER BY T.Name"

    Private connection As IDbConnection

    ''' <summary>
    ''' Имя сервера баз данных
    ''' </summary>
    Public Property ServerName As String

    ''' <summary>
    ''' Включаемые в дамп таблицы
    ''' </summary>
    Public Property IncludeTables As String()

    ''' <summary>
    ''' Имя базы данных
    ''' </summary>
    Public Property Name As String

    ''' <summary>
    ''' Таблицы базы данных
    ''' </summary>
    Public Property Tables As List(Of SQLTable)

    ''' <summary>
    ''' Размер пакета данных (в строках)
    ''' </summary>
    Public Property DataPacketSize As Integer

    ''' <summary>
    ''' Событие начала экспорта таблицы
    ''' </summary>
    Public Event TableExporting As EventHandler(Of TableExportEventArgs)

    ''' <summary>
    ''' Установить соединение с сервером баз данных
    ''' </summary>
    Public Sub Connect()
        Dim builder As New SqlConnectionStringBuilder()
        builder.IntegratedSecurity = True
        builder.InitialCatalog = Name
        builder.DataSource = ServerName
        connection = New SqlConnection(builder.ConnectionString)
        connection.Open()
    End Sub

    ''' <summary>
    ''' Анализировать структуру таблиц
    ''' </summary>
    Public Sub AnalyzeTables()
        For Each table In FetchTables()
            Tables.Add(table)
        Next
    End Sub

    ''' <summary>
    ''' Выполнить экспорт данных в поток
    ''' </summary>
    ''' <param name="outStream">Поток, куда будут экспотрированы данные</param>
    Public Sub Export(outStream As Stream)
        Dim writer As New StreamWriter(outStream)
        Dim fullName = String.Concat(ServerName, "\", Name)
        writer.WriteLine("-- Скрипт экспорта данных БД {0}", fullName)
        writer.WriteLine()
        writer.WriteLine("SET NOCOUNT ON")
        writer.WriteLine("SET DATEFORMAT 'dmy'")
        writer.WriteLine("GO")
        writer.WriteLine()
        writer.WriteLine("USE [{0}]", Name)
        writer.WriteLine()
        writer.Flush()

        For Each table In Tables
            RaiseEvent TableExporting(Me, New TableExportEventArgs(table.Name))
            table.ExportData(connection, outStream, DataPacketSize)
        Next

        writer.WriteLine()
        writer.WriteLine("GO")
        writer.Flush()
    End Sub

    ''' <summary>
    ''' Закрывает соединение с базой данных
    ''' </summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        If connection IsNot Nothing AndAlso _
            connection.State = ConnectionState.Open Then
            connection.Close()
        End If
    End Sub

    ''' <summary>
    ''' Выдает названия таблиц базы данных
    ''' </summary>
    Private Iterator Function FetchTables() As IEnumerable(Of SQLTable)
        Dim namesCmd = connection.CreateCommand()
        namesCmd.CommandType = CommandType.Text
        namesCmd.CommandText = SchemaSql
        Dim reader As IDataReader = namesCmd.ExecuteReader()

        Dim currentTable As SQLTable = Nothing
        Dim currentTableName As String = String.Empty
        Dim row(2) As String

        Dim tableFilterApplied = IncludeTables IsNot Nothing _
            AndAlso IncludeTables.Length > 0

        While reader.Read()
            reader.GetValues(row)
            Dim name As String = row(0)
            If tableFilterApplied Then
                If Not IncludeTables.Contains(name) Then
                    Continue While
                End If
            End If

            If name <> currentTableName Then
                currentTableName = name
                If currentTable IsNot Nothing Then
                    Yield currentTable
                End If
                currentTable = New SQLTable(name)
            End If
            currentTable.AppendDefinition(row)
        End While
        reader.Close()

        Yield currentTable
    End Function

    Public Sub New(serverName As String, dbName As String)
        Me.ServerName = serverName
        Me.Name = dbName
        Me.Tables = New List(Of SQLTable)
    End Sub

End Class
