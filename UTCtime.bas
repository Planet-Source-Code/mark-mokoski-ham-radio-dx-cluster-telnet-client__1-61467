Attribute VB_Name = "UTC_time"
    Option Explicit
    '****************************************************************************
    '
    '   Set of functions to return UTC Time / Date using some API calls
    '   Format of functions similar to standard Time / Date
    '   funtions in VB (Time, Time$, Date, Date$, Now)
    '
    '   Date functions use TimeSerial, DateSerial and Named Formats in VB,
    '   so the output of date functions will follow conventions for the users
    '   country/regional settings.
    '
    '   Example: United States date is MM/DD/YYYY, Europe is DD/MM/YYYY.
    '
    '   Function examples are in United States format
    '
    '   Also there are two funtions that return the name of the
    '   time zone the system is set for, one the short name (like "EST"),
    '   the second for the long name (like "Eastern Standard Time")
    '   Some of the time zone names don't realy work
    '   well with short names.  But it works fine for most
    '   U.S. and Canada time zones.  Change the code as you see fit.
    '   Short and Long Time Zone Names change to "Daylight Time"
    '   if a daylight time zone is selected.
    '
    '   List of new functions with there VB local time equivalent
    '   UTC_time function   VB function     Format
    '   UTCtime             Time$           24 Hour HH:MM:SS
    '   UTCtime2            Time            Local format
    '   UTCdate             Date            MM/DD/YYYY  (region dependent)
    '   UTCnow              Now             MM/DD/YYYY HH:MM:SS AM/PM   (region dependent)
    '   UTCoffset           ---             Current Offset in Minutes from UTC
    '   shortTZname         ----            XXX Ex: "EST", 3 to 5 letters
    '   longTZname          ----            Long name Ex: "Eastern Standard Time"
    '   ISOdate             Date            ISO 8601 format yyyy-mm-dd
    '   ISOtime             Time            ISO 8601 format hh:mm:ssZ
    '   ISOnow              Now             ISO 8601 format yyyy-mm-ddThh:mm:ssZ
    '
    '   Time zone and UTC time functions assume the correct time zone is selected
    '   in the Time/Date properties and clock set correct local time.
    '
    '   Mark Mokoski, 03-JAN-2003
    '
    '***************************************************************************

    Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SystemTime)
    Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long


        Private Type SystemTime
            wYear                                                             As Integer
            wMonth                                                            As Integer
            wDayOfWeek                                                        As Integer
            wDay                                                              As Integer
            wHour                                                             As Integer
            wMinute                                                           As Integer
            wSecond                                                           As Integer
            wMilliseconds                                                     As Integer
        End Type

        Private Type TIME_ZONE_INFORMATION
            Bias                                                              As Long
            StandardName(32)                                                  As Integer
            'StandardName(32)                                                 As String
            StandardDate                                                      As SystemTime
            StandardBias                                                      As Long
            DaylightName(32)                                                  As Integer
            'DaylightName(32)                                                 As String
            DaylightDate                                                      As SystemTime
            DaylightBias                                                      As Long
        End Type

    Dim sysTime                                                               As SystemTime
    Dim TZinfo                                                                As TIME_ZONE_INFORMATION

Public Function UTCtime()

    'Format: HH:MM:SS in 24 hour format (like time$ function)

    Call GetSystemTime(sysTime)
   
    UTCtime = Format(Str$(sysTime.wHour), "00") & ":" & Format(Str$(sysTime.wMinute), "00") & ":" & Format(Str$(sysTime.wSecond), "00")
    
End Function

Public Function UTCtime2()

    'Format: HH:MM:SS in local format (like time function)

    Call GetSystemTime(sysTime)
    
    UTCtime2 = TimeSerial(sysTime.wHour, sysTime.wMinute, sysTime.wSecond)
    
End Function

Public Function UTCnow()

    'Format: Like "Now" function. ex: "1/2/2003 4:35:15 PM"
    'Functions returns format above (region dependent)

    Call GetSystemTime(sysTime)
    
    UTCnow = DateSerial(sysTime.wYear, sysTime.wMonth, sysTime.wDay) & " " & TimeSerial(sysTime.wHour, sysTime.wMinute, sysTime.wSecond)
       
    '****Can also show as MEDIUM DATE. Comment the line above and uncomment below
    'UTCnow = UCase(Format(DateSerial(sysTime.wYear, sysTime.wMonth, sysTime.wDay), "medium date")) & " " & TimeSerial(sysTime.wHour, sysTime.wMinute, sysTime.wSecond)
    
    '****Can also show as LONG DATE. Comment the lines above and uncomment below
    'UTCnow = Format(DateSerial(sysTime.wYear, sysTime.wMonth, sysTime.wDay), "long date") & " " & TimeSerial(sysTime.wHour, sysTime.wMinute, sysTime.wSecond)
    
End Function

Public Function UTCdate()

    'Format: MM/DD/YY ex: 1/1/2002 (like date function)
    'Functions returns format above (region dependent)

    Call GetSystemTime(sysTime)

    UTCdate = DateSerial(sysTime.wYear, sysTime.wMonth, sysTime.wDay)
    
End Function

Public Function UTCoffset()

    'Gets current Bias from UTC in Hours

    Dim TZResult            As Long

    TZResult = GetTimeZoneInformation(TZinfo)
    UTCoffset = Not (TZinfo.Bias - 1)

End Function

Public Function ISOdate()

    'Format: YYYY-MM-DD ex: 2003-01-03
    'Functions returns format above (Fixed ISO Format)

    Call GetSystemTime(sysTime)

    ISOdate = Format(Str(sysTime.wYear), "0000") & "-" & Format(Str(sysTime.wMonth), "00") & "-" & Format(Str(sysTime.wDay), "00")
    
End Function

Public Function ISOtime()

    'Format: hh:mm:ssZ ex: 15:03:47Z ("Z"denotes ZULU or UTC time)
    'Function returns format above (Fixed ISO Format)

    Call GetSystemTime(sysTime)

    ISOtime = Format(Str(sysTime.wHour), "00") & ":" & Format(Str(sysTime.wMinute), "00") & ":" & Format(Str(sysTime.wSecond), "00") & "Z"
    
End Function

Public Function ISOnow()

    'Format: yyyy-mm-ddThh:mm:ssZ ex: 2003-01-03T15:03:47Z
    'Function returns format above (Fixed ISO Format)

    ISOnow = ISOdate & "T" & ISOtime
    
End Function

Public Function shortTZname()

    'Format: XYZ ex: "EST" for Eastern Standard Time
    'Some of the time zone names don't realy work
    'well with short names.  But it works fine for most
    'U.S. and Canada time zones.  Change the code as you see fit.

    Dim m                   As Integer
    Dim TZname              As String

    TZname = longTZname

    'Get first letter of each word of long time zone name to get short name
    'Fist letter, first word
    shortTZname = Mid(TZname, 1, 1)
    'Set pointer for second word
    m = InStr((m + 1), TZname, " ")
    'Loop thru the long time zone name and find first letter of remaining all words

        Do Until m = 0
            shortTZname = shortTZname & Mid(TZname, (m + 1), 1)
            m = InStr((m + 1), TZname, " ")
            DoEvents
        Loop
    
    'Force uppercase for display. It should be already, but just incase
    shortTZname = UCase(shortTZname)
    
End Function

Public Function longTZname()

    'Format: Long Time Zone Name ex: "Eastern Standard Time"

    Dim TZResult            As Long
    Dim i                   As Long
    Dim tempname            As String
    
    TZResult = GetTimeZoneInformation(TZinfo)
    
    'Extract Time Zone Name from returned API call

        For i = 0 To 31

                If TZinfo.StandardName(i) = 0 Then Exit For
            longTZname = longTZname & Chr(TZinfo.StandardName(i))
            DoEvents
        Next

        Select Case TZResult
            Case 0, 1 'Use standard time name
            
            Case 2 'Use daylight savings time name
                tempname = Mid(longTZname, 1, (InStr(1, longTZname, " ")))
                longTZname = tempname & "Daylight Time"
                
        End Select
    
    
    'Trim any spaces in longTZname. Shoud be free of spaces, but just incase
    longTZname = Trim(longTZname)
    
End Function
