Attribute VB_Name = "Kenwood"

    '*************************************************************************************
    '
    'Command format to change frequency, VFO and mode for Kenwood radios
    'Known to work on models TS570(S&D),TS870S,TS2000(S&X and B2000)
    '
    'FA(freq) & FB(freq) command to set frequency of VFO A or B
    'FR(x) & FT(x) command to set RCV and XMIT VFO (0= VFO A, 1=VFO B)
    'MD(x) command to set mode (1=LSB, 2=USB, 3=CW, 4=FM, 5=AM)
    '
    '*************************************************************************************

Public Sub Set_Radio_Kenwood(FreqVal)

    'For Kenwood radios
    'Known to work on TS570(S&D),TS870S,TS2000(S&X and B2000)
    'Parse frequency an VFO data into string to be sent to radio
    'FreqVal = frequency in Hz (11 digits)

    'Convert Frequency in Hz to 11 digit string for Kenwood Radios
    FreqStr = Format(FreqVal, "00000000000")

    'Select the VFO from Properties value and build output string to radio

        Select Case Mid(RadioVFO, 5, 1)
    
            Case "A"    'To send frequency to VFO A
                RadioOutString = "FA" & FreqStr & ";" & "FR0;" & "FT0;"
            Case "B"    'To send frequency to VFO B
                RadioOutString = "FB" & FreqStr & ";" & "FR1;" & "FT1;"
        End Select

    'Find the operating mode based on Band/Mode list in Properties
    'If "NONE returned by function BandMode, do not change current radio mode
    OperMode = BandMode(FreqVal)

        Select Case OperMode

            Case "LSB"
                RadioOutString = RadioOutString & "MD1;"
            Case "USB"
                RadioOutString = RadioOutString & "MD2;"
            Case "CW"
                RadioOutString = RadioOutString & "MD3;"
            Case "FM"
                RadioOutString = RadioOutString & "MD4;"
            Case "AM"
                RadioOutString = RadioOutString & "MD5;"
            Case Else
            
        End Select

    'RadioOutString will look like this for VFO B, 3818 KHz, LSB
    'FB00003818000;FR1;FT1;MD1;

End Sub

Public Function Read_Radio_Kenwood()

    'Return string that send information from radio
    Read_Radio_Kenwood = "IF;"
    'Set length of data to be recieved from radio
    RCV_Data_Len = 38

End Function

Public Function Get_Freq_Kenwood(InfoString)

    'Parse string returned from radio and return value in KHz (XXXX.XX)
    FreqHz = Mid(InfoString, 3, 11)
    Get_Freq_Kenwood = Format((FreqHz / 1000), "########.00")

End Function
