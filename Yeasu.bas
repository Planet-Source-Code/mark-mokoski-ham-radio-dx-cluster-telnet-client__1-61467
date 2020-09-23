Attribute VB_Name = "Yaesu"
    Option Explicit


Public Sub Set_Radio_Yaesu(FreqVal)

End Sub

Public Function Read_Radio_Yaesu()

    'Return string that send information from radio
    'Yaesu radios need 5 byte HEX input, last byte is command "03", get RCV status
    Read_Radio_Yaesu = Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H3)
    'Set length of data to be recieved from radio (in bytes)
    RCV_Data_Len = 5

End Function

Public Function Get_Freq_Yaesu(InfoString)

    'Yaesu radios return a 5 byte HEX output, frequency is BCD encoded in that output
    'Byte 1 = 100 & 10 MHz, byte 2 = 1 MHz & 100KHZ, byte 3 = 10 & 1 KHZ, byte 4 = 100 & 10 Hz, byte 5 = Mode (USB, CW etc)
    'We ignor byte 5, dont care about current mode
    
    Dim TEMPtext            As String
    Dim TEMPchr             As String
    Dim TEMPfreq            As Long
    Dim y                   As Integer
    
    TEMPtext = ""

        For y = 1 To (Len(InfoString) - 1)
            TEMPchr = Asc(Mid(InfoString, y, 1))

                If TEMPchr < 17 Then
                    TEMPtext = TEMPtext & "0" & Hex(Asc(Mid(InfoString, y, 1))) & " "
                Else
                    TEMPtext = TEMPtext & Hex(Asc(Mid(InfoString, y, 1))) & " "
                End If

        Next y

    'Convert to KHz from radio data, Yaesu sends Frequncy to the 10Hz resolution
    TEMPfreq = Val(TEMPtext) / 100
    Get_Freq_Yaesu = Format(TEMPfreq, "########.00")

End Function

Public Sub Yaesu_type()

End Sub
