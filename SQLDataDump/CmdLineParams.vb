''' <summary>
''' Параметры командной строки
''' </summary>
Public Class CmdLineParams
    ''' <summary>
    ''' Имя сервера, к которому выполняется подключение
    ''' </summary>
    Public Property ServerName As String

    ''' <summary>
    ''' База данных для выгрузки данных
    ''' </summary>
    Public Property DBName As String

    ''' <summary>
    ''' Путь к файлу экспортированных данных
    ''' </summary>
    Public Property OutputFile As String

    ''' <summary>
    ''' Размер пакета данных (в строках)
    ''' </summary>
    Public Property PacketSize As Integer

    ''' <summary>
    ''' Включаемые в дамп таблицы
    ''' </summary>
    Public Property TablesToInclude As String()

    Public Sub New()
        ServerName = "(local)"
        DBName = "master"
        OutputFile = "dump.sql"
        PacketSize = 0
    End Sub
End Class