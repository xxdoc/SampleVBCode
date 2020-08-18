VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimeZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#####################################################################################
'#  UNIX Epoch Timestamp via API (+ Other Time-Related Subs) (modCnTime.bas)
'#      By: Nick Campbeln
'#
'#      Revision History:
'#          1.1 (Apr 2, 2003):
'#              Added DateSerialToTimestamp() and isDaylightSavings(), completed work on TimestampToDate()
'#              Contributed to PSC.com (Apr 19, 2003)
'#          1.0 (Aug 26, 2002):
'#              Initial Release
'#
'#      Copyright © 2002-2003 Nick Campbeln (opensource@nick.campbeln.com)
'#          This source code is provided 'as-is', without any express or implied warranty. In no event will the author(s) be held liable for any damages arising from the use of this source code. Permission is granted to anyone to use this source code for any purpose, including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:
'#          1. The origin of this source code must not be misrepresented; you must not claim that you wrote the original source code. If you use this source code in a product, an acknowledgment in the product documentation would be appreciated but is not required.
'#          2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original source code.
'#          3. This notice may not be removed or altered from any source distribution.
'#              (NOTE: This license is borrowed from zLib.)
'#
'#  Please remember to vote on PSC.com if you like this code!
'#  Code URL:
'#####################################################################################

    '#### Required API definitions
Private Declare Function GetTimeZoneInformation Lib "kernel32.dll" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Private Declare Sub GetLocalTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)

    '#### Required UDT definitions
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
  Bias As Long
  StandardName(0 To 31) As Integer
  StandardDate As SYSTEMTIME
  StandardBias As Long
  DaylightName(0 To 31) As Integer
  DaylightDate As SYSTEMTIME
  DaylightBias As Long
End Type

    '#### Required Const definitions
Private Const cSecsInStdYear As Long = 31536000
Private Const cSecInDay As Long = 86400



'#####################################################################################
'# Public Subs/Functions
'#####################################################################################
'#########################################################
'# Converts the passed dDateTime into a correctly offset UNIX Epoch Timestamp
'#########################################################
Public Function DateToTimestamp(ByVal dDateTime As Date) As Double
    Dim oTime As SYSTEMTIME

        '#### Setup the oTime UDT with the elements consumed in BuildTimestamp() based on the passed dDate
    With oTime
        .wYear = Year(dDateTime)
        .wMonth = Month(dDateTime)
        .wDay = Day(dDateTime)
        .wHour = Hour(dDateTime)
        .wMinute = Minute(dDateTime)
        .wSecond = Second(dDateTime)
       '.wMilliseconds = 0
    End With

        '#### Toss the oTime UDT into BuildTimestamp() and return the result to the caller
    DateToTimestamp = BuildTimestamp(oTime)
End Function


'#########################################################
'# Converts the passed serialized date into a correctly offset UNIX Epoch Timestamp
'#########################################################
Public Function DateSerialToTimestamp(ByVal iYear As Integer, ByVal iMonth As Integer, ByVal iDay As Integer, ByVal iHour As Integer, ByVal iMinute As Integer, ByVal iSecond As Integer, ByVal iMilliseconds As Integer) As Double
    Dim oTime As SYSTEMTIME

        '#### Setup the oTime UDT with the elements consumed in BuildTimestamp() based on the passed dDate
    With oTime
        .wYear = iYear
        .wMonth = iMonth
        .wDay = iDay
        .wHour = iHour
        .wMinute = iMinute
        .wSecond = iSecond
        .wMilliseconds = iMilliseconds
    End With

        '#### Toss the oTime UDT into BuildTimestamp() and return the result to the caller
    DateSerialToTimestamp = BuildTimestamp(oTime)
End Function


'#########################################################
'# Takes the passed iMonth and iYear and returns the correct number of days for iMonth
'#########################################################
Public Function DaysInMonth(ByVal iMonth As Integer, ByVal iYear As Integer) As Integer
        '#### Determine the passed iMonth and process accordingly
    Select Case iMonth
            '#### If iMonth is Jan, Mar, May, Jul, Aug, Oct or Dec
        Case 1, 3, 5, 7, 8, 10, 12
            DaysInMonth = 31

            '#### If iMonth is Feb
        Case 2
                '#### If iYear is a leap year, return 29
            If (isLeapYear(iYear)) Then
                DaysInMonth = 29

                '#### Else iYear is a standard year, so return 28
            Else
                DaysInMonth = 28
            End If

            '#### If iMonth is invalid
        Case Is < 1, Is > 12
            DaysInMonth = 0

            '#### Else iMonth is Apr, Jun, Sep, or Nov
        Case Else
            DaysInMonth = 30
    End Select
End Function


'#########################################################
'# Calculates the current (ie- daylight savings) time zone offset in hours
'#########################################################
Public Function GetCurrentTimeZoneOffset() As Single
    Dim oLocalTime As SYSTEMTIME
    Dim oGMT As SYSTEMTIME
    Dim iMinuteDiff As Integer
    Dim iHourDiff As Integer

        '#### Call the GetXTime() APIs to collect the system's time information
    Call GetLocalTime(oLocalTime)
    Call GetSystemTime(oGMT)

        '#### If local time is today
    If (oLocalTime.wDay = oGMT.wDay) Then
            '#### Subtract oGMT from oLocalTime
        iHourDiff = oLocalTime.wHour - oGMT.wHour

        '#### Else if local time is yesterday
    ElseIf (oLocalTime.wDay < oGMT.wDay) Then
            '#### Subtract oGMT from oLocalTime less 24 (so a negetive value is returned)
        iHourDiff = (oLocalTime.wHour - 24) - oGMT.wHour

        '#### Else local time is tomorrow
    Else
            '#### Subtract oLocalTime from oGMT less 24
        iHourDiff = oLocalTime.wHour - (oGMT.wHour - 24)
    End If

        '#### Determine the iMinuteDiff
    iMinuteDiff = Gap(CDbl(oGMT.wMinute), CDbl(oLocalTime.wMinute), 60)

        '#### If iMinuteDiff has a value and the current minute is less then iMinuteDiff
    If (iMinuteDiff > 0 And oLocalTime.wMinute < iMinuteDiff) Then
        iHourDiff = iHourDiff - 1
    End If

        '#### Calculate the time difference between the oLocalTime and oGMT (converting partial hours into a decimal hour), returning the result to the caller
    GetCurrentTimeZoneOffset = iHourDiff + (iMinuteDiff / 60)
End Function


'#########################################################
'# Returns the time zone offset in hours as recorded by the system
'#########################################################
Public Function GetTimeZoneOffset() As Single
    Dim oTimeZone As TIME_ZONE_INFORMATION

        '#### Call the API to set oTimeZone to the current system time zone info
    Call GetTimeZoneInformation(oTimeZone)

        '#### Return the Bias element, converted into hours from minutes
    GetTimeZoneOffset = (-1 * (oTimeZone.Bias / 60))
End Function


'#########################################################
'#
'#########################################################
Public Function isDaylightSavings() As Boolean
        '####
    isDaylightSavings = (GetTimeZoneOffset() <> GetCurrentTimeZoneOffset())
End Function


'#########################################################
'# Correctly calculates if the passed iYear is a leap year
'#      The three rules which the Gregorian calendar uses to determine leap year are as follows:
'#          1) Years divisible by four are leap years, unless...
'#          2) Years also divisible by 100 are not leap years, except...
'#          3) Years divisible by 400 are leap years.
'#########################################################
Public Function isLeapYear(ByVal iYear As Integer) As Boolean
        '#### If iYear is evenly divisible by 4, it might be a leap year
    If ((iYear Mod 4) = 0) Then
            '#### If iYear is evenly divisible by 100, it might be a leap year
        If ((iYear Mod 100) = 0) Then
                '#### If iYear is also evenly divisible by 400, it is a leap year (according to rule #3)
            If ((iYear Mod 400) = 0) Then
                isLeapYear = True

                '#### Else iYear is not a leap year (according to rule #2)
            Else
                isLeapYear = False
            End If

            '#### Else iYear must be a leap year (according to rule #1 and #2)
        Else
            isLeapYear = True
        End If

        '#### Else iYear is not a leap year as it is not evenly divisible by 4 (according to rule #1)
    Else
        isLeapYear = False
    End If
End Function


'#########################################################
'# Returns the current local time as a correclty offset UNIX Epoch Timestamp
'#########################################################
Public Function Timestamp() As Double
    Dim oTime As SYSTEMTIME

        '#### Call the API to set oTime to the current system time
    Call GetLocalTime(oTime)

        '#### Toss the oTime UDT into BuildTimestamp() and return the result to the caller
    Timestamp = BuildTimestamp(oTime)
End Function


'#########################################################
'# Converts an offset UNIX Epoch Timestamp into a Date
'#    NOTE: Since CDate() doesn't like milliseconds on the end of its passed string-date, all the lines refering to millisecond calculation have been commented out
'#########################################################
'!'
Public Function TimestampToDate(ByVal dTimestamp As Double) As Date
    Dim oTime As SYSTEMTIME
   'Dim dMilliseconds As Double

        '#### Calculate in the GMT offset (subtract the seconds in an hour times the current time zone offset less 24 hours)
    dTimestamp = dTimestamp + (GetCurrentTimeZoneOffset() * (cSecInDay / 24))

        '#### With the passed oTime structure
    With oTime
            '#### Init wYear to the UNIX Epoch year of 1970
        .wYear = 1970

            '#### If dTimestamp is at or past the UNIX Epoch (January 1 1970 00:00:00 GMT)
        If (dTimestamp >= 0) Then
                '#### Do while dTimestamp is greater then the cSecsInStdYear
            Do While (dTimestamp >= cSecsInStdYear)
                    '#### If wYear is a leap year
                If (isLeapYear(.wYear)) Then
                        '#### If dTimestamp still has as much or more seconds then that in a leap year
                    If (dTimestamp >= (cSecsInStdYear + cSecInDay)) Then
                            '#### Remove one leap year's worth of seconds from dTimestamp
                        dTimestamp = dTimestamp - (cSecsInStdYear + cSecInDay)

                        '#### Else dTimestamp is refering to December 31st in a leap year
                    Else
                            '#### Un-inc wYear and fall out of the loop now
                        .wYear = .wYear - 1
                        Exit Do
                    End If

                    '#### Else wYear is not a leap year
                Else
                        '#### Remove one non-leap year's seconds from dTimestamp
                    dTimestamp = dTimestamp - cSecsInStdYear
                End If

                    '#### Inc wYear for the next loop
                .wYear = .wYear + 1
            Loop

                '#### Borrow the use of wDayOfWeek to hold the "leap year offset", adding one day if wYear is a leap year
            If (isLeapYear(.wYear)) Then .wDayOfWeek = 1 Else .wDayOfWeek = 0

            '#### Else dTimestamp is before the UNIX Epoch
        Else
                '#### Invert dTimestamp
            dTimestamp = (dTimestamp * -1)

                '#### Do while dTimestamp is greater then the cSecsInStdYear
            Do While (dTimestamp >= cSecsInStdYear)
                    '#### If wYear is a leap year
                If (isLeapYear(.wYear)) Then
                        '#### If dTimestamp still has as much or more seconds then that in a leap year
                    If (dTimestamp >= (cSecsInStdYear + cSecInDay)) Then
                            '#### Remove one leap year's worth of seconds from dTimestamp
                        dTimestamp = dTimestamp - (cSecsInStdYear + cSecInDay)

                        '#### Else dTimestamp is refering to December 31st in a leap year
                    Else
                            '#### Un-dec wYear and fall out of the loop now
                        .wYear = .wYear + 1
                        Exit Do
                    End If

                    '#### Else wYear is not a leap year
                Else
                        '#### Remove one non-leap year's seconds from dTimestamp
                    dTimestamp = dTimestamp - cSecsInStdYear
                End If
            
                    '#### Dec wYear for the next loop
                .wYear = .wYear - 1
            Loop

                '#### If the above determined wYear is a leap year
            If (isLeapYear(.wYear)) Then
                    '#### Invert dTimestamp so the pre-Epoch is calculated correctly below (since a negetive epoch counts the days until Dec 31st of it's year, we invert it so that it counts the days since Jan 1 as that's how the calc below is geared)
                dTimestamp = ((dTimestamp - (cSecsInStdYear + cSecInDay)) * -1)

                    '#### Borrow the use of wDayOfWeek to hold the "leap year offset", adding one day if wYear is a leap year
                .wDayOfWeek = 1

                '#### Else wYear is not a leap year
            Else
                    '#### Invert dTimestamp so the pre-Epoch is calculated correctly below (since a negetive epoch counts the days until Dec 31st of it's year, we invert it so that it counts the days since Jan 1 as that's how the calc below is geared)
                dTimestamp = ((dTimestamp - cSecsInStdYear) * -1)

                    '#### Borrow the use of wDayOfWeek to hold the "leap year offset", adding one day if wYear is a leap year
                .wDayOfWeek = 0
            End If
        End If

            '#### Determine the "day of year" and process the month accordingly
        Select Case Fix(dTimestamp / cSecInDay)
                '#### If we are in the month of December
            Case Is >= (334 + .wDayOfWeek)
                    '#### Subtract the seconds making up the month of November from dTimestamp and set wMonth
                dTimestamp = dTimestamp - ((334 + .wDayOfWeek) * cSecInDay)
                .wMonth = 12

                '#### If we are in the month of November
            Case Is >= (304 + .wDayOfWeek)
                    '#### Subtract the seconds making up the month of October from dTimestamp and set wMonth
                dTimestamp = dTimestamp - ((304 + .wDayOfWeek) * cSecInDay)
                .wMonth = 11

                '#### If we are in the month of October
            Case Is >= (273 + .wDayOfWeek)
                    '#### Subtract the seconds making up the month of September from dTimestamp and set wMonth
                dTimestamp = dTimestamp - ((273 + .wDayOfWeek) * cSecInDay)
                .wMonth = 10

                '#### If we are in the month of September
            Case Is >= (242 + .wDayOfWeek)
                    '#### Subtract the seconds making up the month of August from dTimestamp and set wMonth
                dTimestamp = dTimestamp - ((242 + .wDayOfWeek) * cSecInDay)
                .wMonth = 9

                '#### If we are in the month of August
            Case Is >= (212 + .wDayOfWeek)
                    '#### Subtract the seconds making up the month of July from dTimestamp and set wMonth
                dTimestamp = dTimestamp - ((212 + .wDayOfWeek) * cSecInDay)
                .wMonth = 8

                '#### If we are in the month of July
            Case Is >= (181 + .wDayOfWeek)
                    '#### Subtract the seconds making up the month of June from dTimestamp and set wMonth
                dTimestamp = dTimestamp - ((181 + .wDayOfWeek) * cSecInDay)
                .wMonth = 7

                '#### If we are in the month of June
            Case Is >= (151 + .wDayOfWeek)
                    '#### Subtract the seconds making up the month of May from dTimestamp and set wMonth
                dTimestamp = dTimestamp - ((151 + .wDayOfWeek) * cSecInDay)
                .wMonth = 6

                '#### If we are in the month of May
            Case Is >= (120 + .wDayOfWeek)
                    '#### Subtract the seconds making up the month of April from dTimestamp and set wMonth
                dTimestamp = dTimestamp - ((120 + .wDayOfWeek) * cSecInDay)
                .wMonth = 5

                '#### If we are in the month of April
            Case Is >= (90 + .wDayOfWeek)
                    '#### Subtract the seconds making up the month of March from dTimestamp and set wMonth
                dTimestamp = dTimestamp - ((90 + .wDayOfWeek) * cSecInDay)
                .wMonth = 4

                '#### If we are in the month of March
            Case Is >= (59 + .wDayOfWeek)
                    '#### Subtract the seconds making up the month of February from dTimestamp and set wMonth
                dTimestamp = dTimestamp - ((59 + .wDayOfWeek) * cSecInDay)
                .wMonth = 3

                '#### If we are in the month of February
            Case Is >= 31
                    '#### Subtract the seconds making up the month of January from dTimestamp and set wMonth
                dTimestamp = dTimestamp - (31 * cSecInDay)
                .wMonth = 2

                '#### Else we're in the month of January
            Case Else
                    '#### Subtract nothing from dTimestamp as we are in January and set wMonth
               'dTimestamp = dTimestamp - 0
                .wMonth = 1
        End Select

            '#### Since Mod returns a whole number, grab the wMilliseconds (as a whole number) from dTimestamp now (removing them from dTimestamp when we're done)
       '.wMilliseconds = CInt((dTimestamp - Fix(dTimestamp)) * 10000)
        dTimestamp = Fix(dTimestamp)

            '#### Calculate the whole number of wDay (adding 1 for the partial day of today) in dTimestamp then remove them
        .wDay = Fix(dTimestamp / cSecInDay) + 1
        dTimestamp = (dTimestamp Mod cSecInDay)

            '#### Calculate the whole number of wHour (adding 1 for the partial hour of now) in dTimestamp then remove them
        .wHour = Fix(dTimestamp / 3600)
        dTimestamp = (dTimestamp Mod 3600)

            '#### Calculate the whole number of wMinute (adding 1 for the partial minute of now) in dTimestamp then remove them
        .wMinute = Fix(dTimestamp / 60)
        dTimestamp = (dTimestamp Mod 60)

            '#### Calculate the whole number of wSecond (adding 1 for the partial second of now) in dTimestamp then remove them
        .wSecond = Fix(dTimestamp)
        dTimestamp = dTimestamp - .wSecond

            '#### Take the above calculated elements, pass them into the *Serial() functions, then back into CDate() and finially to the caller
       'TimestampToDate = CDate(DateSerial(.wYear, .wMonth, .wDay) & " " & Format(TimeSerial(.wHour, .wMinute, .wSecond), "hh:mm:ss") & "." & .wMilliseconds)
        TimestampToDate = CDate(DateSerial(.wYear, .wMonth, .wDay) & " " & Format(TimeSerial(.wHour, .wMinute, .wSecond), "hh:mm:ss")) ' & "." & .wMilliseconds)
    End With
End Function



'#####################################################################################
'# Private Subs/Functions
'#####################################################################################
'#########################################################
'# Converts the passed oTime UDT into a correctly GMT offset UNIX Epoch Timestamp
'#########################################################
Private Function BuildTimestamp(ByRef oTime As SYSTEMTIME) As Double
    Dim i As Long

        '#### With the passed oTime structure
    With oTime
            '#### Borrow the use of wDayOfWeek to hold the leap year offset, adding one day if wYear is a leap year
        If (isLeapYear(.wYear)) Then .wDayOfWeek = 1 Else .wDayOfWeek = 0

            '#### Determine the wMonth and process accordingly
        Select Case .wMonth
                '#### If it is January, init the return value to 0
            Case 1
                BuildTimestamp = 0 '# (cSecInDay * 0)

                '#### If it is February, init the return value to Jan's days
            Case 2
                BuildTimestamp = (cSecInDay * 31)

                '#### If it is March, init the return value to Jan+Feb's days
            Case 3
                BuildTimestamp = (cSecInDay * (59 + .wDayOfWeek))

                '#### If it is April, init the return value to Jan+Feb+Mar's days
            Case 4
                BuildTimestamp = (cSecInDay * (90 + .wDayOfWeek))

                '#### If it is May, init the return value to Jan+Feb+Mar+Apr's days
            Case 5
                BuildTimestamp = (cSecInDay * (120 + .wDayOfWeek))

                '#### If it is June, init the return value to Jan+Feb+Mar+Apr+May's days
            Case 6
                BuildTimestamp = (cSecInDay * (151 + .wDayOfWeek))

                '#### If it is July, init the return value to Jan+Feb+Mar+Apr+May+Jun's days
            Case 7
                BuildTimestamp = (cSecInDay * (181 + .wDayOfWeek))

                '#### If it is August, init the return value to Jan+Feb+Mar+Apr+May+Jun+Jul's days
            Case 8
                BuildTimestamp = (cSecInDay * (212 + .wDayOfWeek))

                '#### If it is September, init the return value to Jan+Feb+Mar+Apr+May+Jun+Jul+Aug's days
            Case 9
                BuildTimestamp = (cSecInDay * (243 + .wDayOfWeek))

                '#### If it is October, init the return value to Jan+Feb+Mar+Apr+May+Jun+Jul+Aug+Sep's days
            Case 10
                BuildTimestamp = (cSecInDay * (273 + .wDayOfWeek))

                '#### If it is November, init the return value to Jan+Feb+Mar+Apr+May+Jun+Jul+Aug+Sep+Oct's days
            Case 11
                BuildTimestamp = (cSecInDay * (304 + .wDayOfWeek))

                '#### If it is December, init the return value to Jan+Feb+Mar+Apr+May+Jun+Jul+Aug+Sep+Oct+Nov's days
            Case 12
                BuildTimestamp = (cSecInDay * (334 + .wDayOfWeek))
        End Select

            '#### Add in all of the seconds up to the wDay (not including today)
        BuildTimestamp = BuildTimestamp + (cSecInDay * (.wDay - 1))

            '#### Add in all of the seconds up to the wHour
       'BuildTimestamp = BuildTimestamp + (3600 * .wHour)           '# 60 secs * 60 mins * wHour
        BuildTimestamp = BuildTimestamp + ((cSecInDay / 24) * .wHour)

            '#### Add in all of the seconds up to the wMinute
        BuildTimestamp = BuildTimestamp + (60 * .wMinute)           '# 60 secs * wMinute

            '#### Add in the remaining seconds from wSecond
        BuildTimestamp = BuildTimestamp + .wSecond

            '#### Add in the milliseconds from wMilliseconds
        BuildTimestamp = BuildTimestamp + (.wMilliseconds / (10 ^ Len(CStr(.wMilliseconds))))

            '#### If the passed wYear is 1970 or beyond
        If (.wYear >= 1970) Then
                '#### Traverse the years between 1970 and last year (skipping this year as it's not a full year)
            For i = 1970 To (.wYear - 1)
                    '#### If the current year (i) was a leap year
                If (isLeapYear(i)) Then
                    BuildTimestamp = BuildTimestamp + (cSecInDay * 366)

                    '#### Else it was not a leap year
                Else
                    BuildTimestamp = BuildTimestamp + (cSecInDay * 365)
                End If
            Next

            '#### Else the passed wYear is before 1970
        Else
'!'                '#### Invert the current return value by subtracting one non-leap year's worth of seconds (as 1970 was not a leap year)
            BuildTimestamp = BuildTimestamp - (cSecInDay * 365)

                '#### Traverse the years between wYear thru 1969 (1970 - 1) (as the partial year was accounted for above)
            For i = .wYear To 1969
                    '#### If the current year was a leap year
                If (isLeapYear(i)) Then
                    BuildTimestamp = BuildTimestamp - (cSecInDay * 366)

                    '#### Else it was not a leap year
                Else
                    BuildTimestamp = BuildTimestamp - (cSecInDay * 365)
                End If
            Next
        End If

            '#### Calculate in the GMT offset and return the result to the caller (add the seconds in an hour times the current time zone offset less 24 hours)
        BuildTimestamp = BuildTimestamp - (GetCurrentTimeZoneOffset() * (cSecInDay / 24))
    End With
End Function


'#########################################################
'# Determines and returns the gap between dNumber1 and dNumber2 based on lSpan
'#########################################################
Private Function Gap(ByRef dNumber1 As Double, ByRef dNumber2 As Double, ByRef lSpan As Long) As Double
        '#### If dNumber1 is less then dNumber2
    If (dNumber1 < dNumber2) Then
            '#### Return the gap, ensuring its less then lSpan
        Gap = (dNumber2 - dNumber1) Mod lSpan

        '#### Else dNumber1 is more then dNumber2
    Else
            '#### Return the adjusted gap, ensuring its less then lSpan
        Gap = ((lSpan - dNumber1) + dNumber2) Mod lSpan
    End If
End Function

