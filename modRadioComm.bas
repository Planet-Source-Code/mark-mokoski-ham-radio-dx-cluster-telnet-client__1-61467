Attribute VB_Name = "modRadioComm"
    
    Dim DXfreqVal            As Integer

Public Sub SendRadio(DXfreqVal)

    'Find radio type then branch to compile outout string

        If RadioModel = "None" Then
            Exit Sub    'No Radio to control, so exit
        ElseIf Mid(RadioModel, 1, 2) = "TS" Then
            Call Kenwood.Set_Radio_Kenwood(DXfreqVal)   'Kenwood radios supported
        ElseIf Mid(RadioModel, 1, 2) = "IC" Then
            Call ICOM.Set_Radio_ICOM(DXfreqVal)
        ElseIf Mid(RadioModel, 1, 2) = "FT" Then        'FT817D Yaesu supported
            Call Yaesu.Set_Radio_Yaesu(DXfreqVal)
        ElseIf Mid(RadioModel, 1, 6) = "TenTec" Then
            Call TenTec.Set_Radio_TenTec(DXfreqVal)
        End If


    frmTelnet.ComRadio.CommPort = Val(Mid(RadioComPort, 5, 1))
    parity = Mid(RadioComParity, 1, 1)
    'Set up COM Port

        If frmTelnet.ComRadio.PortOpen = True Then
            frmTelnet.ComRadio.PortOpen = False 'Just to make sure !!
        End If

        Select Case RadioComFlow
            Case "None"
                frmTelnet.ComRadio.Handshaking = 0
            Case "XON/XOFF"
                frmTelnet.ComRadio.Handshaking = 1
            Case "RTS/CTS"
                frmTelnet.ComRadio.Handshaking = 2
        End Select

    On Error GoTo ComPortErr
    'If error opening port, display message box warning
    frmTelnet.ComRadio.Settings = RadioComSpeed & "," & parity & "," & RadioComBits & RadioComStop
    frmTelnet.ComRadio.PortOpen = True
    frmTelnet.ComRadio.Output = RadioOutString
    'Wait until all RadioOutString is sent to radio

        Do Until frmTelnet.ComRadio.OutBufferCount = 0
            DoEvents
        Loop

    frmTelnet.ComRadio.PortOpen = False

    Exit Sub
    'If Com Port error, display message box warning
ComPortErr:
    MsgBox "ERROR Opening Commuications Port" & vbCrLf & _
    "Port may be in use by another application or" & vbCrLf & _
    "is not a installed COM port." & vbCrLf & _
    "Check Properties for selected COM port", vbExclamation, "Communications Port ERROR"

End Sub

Public Function RadioFreq()

    'Find radio type then branch to compile outout string

        If RadioModel = "None" Then
            Exit Function    'No Radio to control, so exit
        ElseIf Mid(RadioModel, 1, 2) = "TS" Then
            RadioOutString = Kenwood.Read_Radio_Kenwood  'Kenwood radios supported
        ElseIf Mid(RadioModel, 1, 2) = "IC" Then
            RadioOutString = ICOM.Read_Radio_ICOM
        ElseIf Mid(RadioModel, 1, 2) = "FT" Then        'FT817D Yaesu supported
            RadioOutString = Yaesu.Read_Radio_Yaesu
        ElseIf Mid(RadioModel, 1, 6) = "TenTec" Then
            RadioOutString = TenTec.Read_Radio_TenTec
        End If

    frmTelnet.ComRadio.CommPort = Val(Mid(RadioComPort, 5, 1))
    parity = Mid(RadioComParity, 1, 1)
    'Set up COM Port

        If frmTelnet.ComRadio.PortOpen = True Then
            frmTelnet.ComRadio.PortOpen = False 'Just to make sure !!
        End If

        Select Case RadioComFlow
            Case "None"
                frmTelnet.ComRadio.Handshaking = 0
            Case "XON/XOFF"
                frmTelnet.ComRadio.Handshaking = 1
            Case "RTS/CTS"
                frmTelnet.ComRadio.Handshaking = 2
        End Select

    On Error GoTo ComPortErr2
    'If error opening port, display message box warning
    frmTelnet.ComRadio.Settings = RadioComSpeed & "," & parity & "," & RadioComBits & RadioComStop
    frmTelnet.ComRadio.InputLen = 1
    frmTelnet.ComRadio.PortOpen = True
    'Read the COM port to clean out any crap
    frmTelnet.ComRadio.InputLen = frmTelnet.ComRadio.InBufferCount
    TrashBuffer = frmTelnet.ComRadio.Input
    frmTelnet.ComRadio.Output = RadioOutString
    'Wait until all RadioOutString is sent to radio

        Do Until frmTelnet.ComRadio.OutBufferCount = 0
            DoEvents
        Loop

    'Get responce from radio

    'Clear the buffer
    RadioResponce = ""
    'Wait while radio thinks about sending data

        Do Until frmTelnet.ComRadio.InBufferCount = RCV_Data_Len
            DoEvents
        Loop
  
    'Now that we have something, shift it into the buffer

        Do Until frmTelnet.ComRadio.InBufferCount = 0
            RadioResponce = RadioResponce & frmTelnet.ComRadio.Input
            DoEvents
        Loop

    'Close the COM port
    frmTelnet.ComRadio.PortOpen = False

    'Convert responce string into what we need (freq in KHz, format XXXX.XX)

        If RadioModel = "None" Then
            Exit Function    'No Radio to control, so exit
        ElseIf Mid(RadioModel, 1, 2) = "TS" Then
            RadioFreq = Kenwood.Get_Freq_Kenwood(RadioResponce)  'Kenwood radios supported
        ElseIf Mid(RadioModel, 1, 2) = "IC" Then
            RadioFreq = ICOM.Get_Freq_ICOM(RadioResponce)
        ElseIf Mid(RadioModel, 1, 2) = "FT" Then        'FT817D Yaesu supported
            RadioFreq = Yaesu.Get_Freq_Yaesu(RadioResponce)
        ElseIf Mid(RadioModel, 1, 6) = "TenTec" Then
            RadioFreq = TenTec.Get_Freq_TenTec(RadioResponce)
        End If

    Exit Function
    'If Com Port error, display message box warning
ComPortErr2:
    MsgBox "ERROR Opening Commuications Port" & vbCrLf & _
    "Port may be in use by another application or" & vbCrLf & _
    "is not a installed COM port." & vbCrLf & _
    "Check Properties for selected COM port", vbExclamation, "Communications Port ERROR"

End Function
