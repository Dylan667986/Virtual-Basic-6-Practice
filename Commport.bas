Attribute VB_Name = "PowerController"
Option Explicit

' Basic VB6 helper module for controlling a programmable power supply via USB-to-COM
' Requires a Form or Class that hosts an MSComm control named and passed into these helpers.

Private Const DefaultBaud As Long = 9600          ' Default baud rate for most bench supplies
Private Const DefaultParity As String = "n"       ' No parity
Private Const DefaultDatabits As Long = 8         ' 8 data bits
Private Const DefaultStopbits As Long = 1         ' 1 stop bit
Private Const DefaultTimeoutMs As Long = 2000     ' Response wait timeout in milliseconds
' VB6 sometimes loses the MSComm constants; define the text mode value explicitly to avoid
' red-ink syntax errors when the control library is not referenced yet.
Private Const InputModeText As Integer = 0        ' Equivalent to comInputModeText

Public Sub InitializePort(ByRef comm As MSComm, ByVal comNumber As Integer, _
                          Optional ByVal baud As Long = DefaultBaud)
    ' Ensure the port is closed before applying settings to avoid runtime errors
    If comm.PortOpen Then comm.PortOpen = False

    ' Configure COM port number and serial settings
    comm.CommPort = comNumber
    comm.Settings = CStr(baud) & "," & DefaultParity & "," & DefaultDatabits & "," & DefaultStopbits
    ' Use text mode for SCPI-style ASCII commands
    comm.InputMode = InputModeText
    comm.RThreshold = 1
    comm.InputLen = 0
    comm.SThreshold = 0
    comm.InBufferSize = 1024
    comm.OutBufferSize = 1024
    ' Open the port after configuration
    comm.PortOpen = True
End Sub

Public Sub ClosePort(ByRef comm As MSComm)
    ' Close the port when finished or before reconfiguration
    If comm.PortOpen Then comm.PortOpen = False
End Sub

Public Function SendCommand(ByRef comm As MSComm, ByVal commandText As String, _
                            Optional ByVal waitForReply As Boolean = False) As String
    ' Append CRLF because many supplies expect both characters as a terminator
    Dim payload As String
    payload = commandText & vbCrLf

    ' Write the outbound command to the COM port
    comm.Output = payload
    If Not waitForReply Then Exit Function

    Dim startTick As Single
    startTick = Timer

    Do While (Timer - startTick) * 1000 < DefaultTimeoutMs
        DoEvents
        ' Read the whole buffer if data is available, then return to caller
        If comm.InBufferCount > 0 Then
            SendCommand = comm.Input
            Exit Function
        End If
    Loop

    Err.Raise vbObjectError + 100, "PowerController.SendCommand", "Timeout waiting for reply"
End Function

Public Sub SetVoltage(ByRef comm As MSComm, ByVal volts As Double)
    ' Format voltage to three decimals and send the SCPI VOLT command
    SendCommand comm, "VOLT " & Format$(volts, "0.000")
End Sub

Public Sub SetCurrent(ByRef comm As MSComm, ByVal amps As Double)
    ' Format current to three decimals and send the SCPI CURR command
    SendCommand comm, "CURR " & Format$(amps, "0.000")
End Sub

Public Sub EnableOutput(ByRef comm As MSComm)
    ' Turn on the output relay
    SendCommand comm, "OUTP ON"
End Sub

Public Sub DisableOutput(ByRef comm As MSComm)
    ' Turn off the output relay
    SendCommand comm, "OUTP OFF"
End Sub

Public Function ReadMeasuredVoltage(ByRef comm As MSComm) As Double
    ' Query measured voltage; Val() converts the numeric text into a Double
    Dim response As String
    response = SendCommand(comm, "MEAS:VOLT?", True)
    ReadMeasuredVoltage = Val(response)
End Function

Public Function ReadMeasuredCurrent(ByRef comm As MSComm) As Double
    ' Query measured current; Val() converts the numeric text into a Double
    Dim response As String
    response = SendCommand(comm, "MEAS:CURR?", True)
    ReadMeasuredCurrent = Val(response)
End Function

Public Function QueryIdentification(ByRef comm As MSComm) As String
    ' Return the identification string for logging or verification
    QueryIdentification = SendCommand(comm, "*IDN?", True)
End Function
