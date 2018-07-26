' Use this code inside a project created with the Visual Basic > Windows Desktop > Console Application template.
' Replace the default code in Module1.vb with this code. Then right click the project in Solution Explorer,
' select Properties, and set the Startup Object to PortChat.

Imports System.Data.Odbc
Imports System.IO.Ports
Imports System.Threading

Public Class PortChat
    Shared _continue As Boolean
    Shared _serialPort As SerialPort
    Shared conn As OdbcConnection
    Shared cmd = New OdbcCommand("", conn)
    Shared SQLCommandText As String = ""
    Shared _Debug As Boolean

    Public Shared Sub Main()
Start:

        Try

            SQLCommandText = My.MySettings.Default.SQLCommandText

            If My.MySettings.Default.Debug = "On" Then
                _Debug = True
            Else
                _Debug = False
            End If

            conn = New OdbcConnection

            conn.ConnectionString = My.MySettings.Default.constr
            conn.ConnectionTimeout = CInt(My.MySettings.Default.ConnectionTimeout)
            conn.Open()

            cmd = New OdbcCommand("", conn)

            cmd.CommandType = CommandType.Text
            cmd.CommandTimeout = CInt(My.MySettings.Default.CommandTimeout)

            Dim portName As String
            Dim message As String
            Dim stringComparer__1 As StringComparer = StringComparer.OrdinalIgnoreCase
            Dim readThread As New Thread(AddressOf Read)

            ' Create a new SerialPort object with default settings.
            _serialPort = New SerialPort()

            _serialPort.PortName = My.MySettings.Default.PortName
            _serialPort.BaudRate = CInt(My.MySettings.Default.BaudRate)
            _serialPort.Parity = CType([Enum].Parse(GetType(Parity), My.MySettings.Default.Parity, True), Parity)
            _serialPort.DataBits = CInt(My.MySettings.Default.DataBits)
            _serialPort.StopBits = CType([Enum].Parse(GetType(StopBits), My.MySettings.Default.StopBits, True), StopBits)
            _serialPort.Handshake = CType([Enum].Parse(GetType(Handshake), My.MySettings.Default.Handshake, True), Handshake)

            ' Set the read/write timeouts
            _serialPort.ReadTimeout = (CInt(My.MySettings.Default.CommandTimeout) * 1000) + 500
            _serialPort.WriteTimeout = 50

            _serialPort.Open()
            _continue = True

            readThread.Start()

            portName = My.MySettings.Default.PortName

            Console.WriteLine("Type QUIT to exit")

            While _continue
                message = Console.ReadLine()

                If stringComparer__1.Equals("quit", message) Then
                    _continue = False
                Else
                    _serialPort.WriteLine([String].Format("<{0}>: {1}", portName, message))
                End If
            End While

            readThread.Join()
            _serialPort.Close()

            Try
                If conn.State <> ConnectionState.Closed Then conn.Close()
            Catch ex2 As Exception

            End Try

        Catch ex As Exception

            If _Debug Then
                Console.WriteLine("")
                Console.WriteLine("")
                Console.WriteLine("")
                Console.WriteLine("Error:" + ex.Message)
                Console.WriteLine("")
                Console.WriteLine("")
                Console.WriteLine("")
                Console.WriteLine("Press key continue....")
                Console.ReadKey()
                Console.Clear()
            End If

            Try
                If conn.State <> ConnectionState.Closed Then conn.Close()
            Catch ex2 As Exception

            End Try

            GoTo Start

        End Try

    End Sub

    Public Shared Sub Read()
        While _continue
            Try
                Dim message As String = _serialPort.ReadLine()

                If message.StartsWith("**:") Then
                    If _Debug Then Console.WriteLine("message.Contains('**:')")

                    Dim value As String = message
                    value = value.Substring(4, 8).Trim.Replace(",", ".")
                    If _Debug Then Console.WriteLine("value:" + value)

                    If IsNumeric(value) Then
                        cmd.CommandText = String.Format("UPDATE peso SET peso_atual = {0};", value)
                        cmd.ExecuteNonQuery()
                        If _Debug Then Console.WriteLine("cmd.ExecuteNonQuery()")
                    End If
                End If

                Console.WriteLine(message)

            Catch ex As TimeoutException
                Console.WriteLine("Read1:" + ex.Message)
            Catch ex As Exception
                Console.WriteLine("Read2:" + ex.Message)
            End Try
        End While
    End Sub

End Class