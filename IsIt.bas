'------------------------------------------------------------------------------
'Purpose  : Evaluates a passed date/time condition and returns an
'           ERRORLEVEL, if the condition has been met.
'           Condition has been met: ERRORLEVEL = 1
'           Condition has NOT been met: ERRORLEVEL = 0
'           Invalid parameter: ERRORLEVEL = 254
'           Other error during program execution: ERRORLEVEL = 255
'
'Prereq.  : baCmdLine.sll
'Note     : -
'
'   Author: Knuth Konrad 2018
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
#Compile Exe ".\IsIt.exe"
#Option Version5
#Dim All

#Link "baCmdLine.sll"

#Break On
#Debug Error On
#Tools On

%VERSION_MAJOR = 1
%VERSION_MINOR = 0
%VERSION_REVISION = 0

' Version Resource information
#Include ".\IsItRes.inc"
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
' Comparison unit
%UNIT_OTHER = 0

%UNIT_DATE = 1             ' Full date as (yyyymmdd)
%UNIT_DATE_YEAR = 2        ' Year only (yyyy)
%UNIT_DATE_MONTH = 3       ' Month only (mm)
%UNIT_DATE_DAY = 4         ' Day only (dd)
%UNIT_DATE_YEARMONTH = 5   ' Year and month (yyyymm)
%UNIT_DATE_MONTHDAY = 6    ' Month and date  (mmdd)

%UNIT_TIME = 7             ' Full time as (hhnnss)
%UNIT_TIME_HOUR = 8        ' Hour only (hh)
%UNIT_TIME_MINUTE = 9      ' Minute only (nn)
%UNIT_TIME_SECOND = 10     ' Second only (ss)
%UNIT_TIME_HOURMINUTE = 11 ' Hour and minute (hhnn)

%UNIT_DAY_WEEK = 12        ' Day of week 0-6 = Sunday-Saturday
%UNIT_DAY_FDW = 13         ' First day of week

' Return value
%RET_INVALID_ARG = 254
%RET_OTHER_ERROR = 255

' Console colors
%Green = 2
%Red = 4
%White = 7
%Yellow = 14
%LITE_GREEN = 10
%LITE_RED = 12
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
   ' Valid CLI parameters are:
   ' /c or /compare
   ' /u= or /unit
   ' /v= or /value
   ' /fdw or /firstdayofweek
' Parameters passed via CLI
Type CfgTYPE
   Compare As Long
   Unit As Long
   FDW As Long
End Type

' Weekdays
Enum Weekday
   Sunday = 0
   Monday = 1
   Tuesday = 2
   Wednesday = 3
   Thursday = 4
   Friday = 5
   Saturday = 6
End Enum
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
#Include Once "win32api.inc"
'#Include "ImageHlp.inc"
#Include "sautilcc.inc"       ' General console helpers
'#Include "IbaCmdLine.inc"
'------------------------------------------------------------------------------
'*** Variabels ***
'------------------------------------------------------------------------------
'==============================================================================

Function PBMain () As Long
'------------------------------------------------------------------------------
'Purpose  : Programm startup method
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sUnit, sValue, sCmd, sTemp As String
   Local i As Dword
   Local lUnit, lFDW, lResult As Long
   Local vntResult As Variant
   Local udtCfg As CfgTYPE

   Local oPTNow As IPowerTime
   Let oPTNow = Class "PowerTime"

   ' Application intro
   ConHeadline "IsIt", %VERSION_MAJOR, %VERSION_MINOR, %VERSION_REVISION
   ConCopyright "2022", $COMPANY_NAME
   Print ""

   Trace New ".\IsIt.tra"

   ' *** Parse the parameters
   ' Initialization and basic checks
   sCmd = Command$

   Local o As IBACmdLine
   Local vnt As Variant

   Let o = Class "cBACmdLine"

   If IsFalse(o.Init(sCmd)) Then
      Print "Couldn't parse parameters: " & sCmd
      Print "Type IsIt /? for help"
      Let o = Nothing
      PBMain = %RET_INVALID_ARG
      Exit Function
   End If

   If Len(Trim$(Command$)) < 1 Or InStr(Command$, "/?") > 0 Then
      ShowHelp
      PBMain = %RET_INVALID_ARG
      Exit Function
   End If

   ' Parse the passed parameters
   ' Valid CLI parameters are:
   ' /c or /compare
   ' /u= or /unit
   ' /v= or /value
   ' /fdw or /firstdayofweek
   i = o.ValuesCount

   If i < 3 Then
      Print "Invalid number of parameters."
      Print ""
      ShowHelp
      PBMain = %RET_INVALID_ARG
      Exit Function
   End If

   Trace On

   ' Parse CLI parameters

   ' ** Comparison
   If IsTrue(o.HasParam("c", "compare")) Then
      vntResult = o.GetValueByName("c", "compare")
      udtCfg.Compare = Val(Variant$(vntResult))
   End If

   ' ** Unit
   If IsTrue(o.HasParam("u", "unit")) Then
      sTemp = Variant$(o.GetValueByName("u", "unit"))
      sUnit = LCase$(Trim$(Remove$(sTemp, $Dq)))
   End If

   ' ** Value
   If IsTrue(o.HasParam("v", "value")) Then
      sTemp = Variant$(o.GetValueByName("v", "value"))
      sValue = Trim$(Remove$(sTemp, $Dq))
   End If

   ' ** First day of week, if omitted, defaults to Monday (1)
   If IsTrue(o.HasParam("fdw", "firstdayofweek")) Then
      vntResult = o.GetValueByName("fdw", "firstdayofweek")
      udtCfg.FDW = Val(Variant$(vntResult))
   Else
      udtCfg.FDW = %Weekday.Monday
   End If

   ' *** Echo the CLI parameters

   ' Valid CLI parameters are:
   ' /c or /compare
   ' /u= or /unit
   ' /v= or /value
   ' /fdw or /firstdayofweek
   StdOutEx "Command line arguments", sCmd
   StdOutEx "Comparison", Format$(udtCfg.Compare) & " (" & GetCompareStr(udtCfg.Compare) & ")"
   StdOutEx "Unit", sUnit & " (" & GetUnit(sUnit, lUnit) & ")"
   StdOutEx "Compare to", sValue
   udtCfg.Unit = lUnit

   ' Determine the weekday, lUnit holds the number (0-6), the method returns the string ("Sunday"-"Saturday")
   sTemp = GetUnit("fdw", lUnit, udtCfg.FDW)
   StdOutEx "First day o. week", Format$(udtCfg.FDW) & " (" & sTemp & ")"
   StdOut ""

   ' *** Sanity checks of CLI parameters
   ' Unit
   If lUnit = %UNIT_OTHER Then
      Print "Invalid unit: " & sUnit
      Print ""
      PBMain = %RET_INVALID_ARG
      ShowHelp
      Exit Function
   End If

   ' Value - depends on unit
   Local oPTValue As IPowerTime
   Let oPTValue = Class "PowerTime"

   ' oPTValue is passed ByRef and will be filled by ValidateValue with
   ' proper values from /value, depending on /unit
   If IsFalse(ValidateValue(udtCfg.Unit, sValue, oPTValue)) Then
      Print "Invalid value for this comparison/unit: " & sValue
      Print ""
      PBMain = %RET_INVALID_ARG
      ShowHelp
      Exit Function
   End If

   ' Show the current date / time depending on the unit
   Call oPTNow.Now()
   StdOutEx "Current date/time", GetCurrentDateTimeStr(oPTNow)

   ' *** Let's compare ...
   Con.StdOut ""

   Try

      lResult = IsIt(oPTNow, oPTValue, sValue, udtCfg)

      Con.StdOut "Done. IsIt: ";
      If IsTrue(lResult) Then
         Con.Color %Green, -1
      Else
         Con.Color %Red, -1
      End If
      Con.StdOut IIf$(IsTrue(lResult), "True", "False")

   Catch

      lResult = %RET_OTHER_ERROR
      Con.StdOut "An error occured: ";
      Con.Color %LITE_RED, -1
      Con.StdOut ErrString(Err)
      ErrClear

   End Try

   Trace Off
   Trace Close

   Con.Color %White, -1
   Con.StdOut ""

   PBMain = lResult

End Function
'---------------------------------------------------------------------------

Function IsIt(ByVal oCurrent As IPowerTime, ByVal oValue As IPowerTime, ByVal sValue As String, _
   ByVal udtCfg As CfgTYPE) As Long
'------------------------------------------------------------------------------
'Purpose  : Does the actual evaluation/comparison of the passed date/time value
'           to the current date/time.
'           unit.
'Prereq.  : -
'Parameter: oCurrent - Current date/time
'           lCompare - Comparison to perform
'           lUnit    - What to compare
'           sValue   - Value against which the comparison takes place
'           oValue   - iPowerTime object filled with sValue, where relevant
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local p As IPowerTime
   Let p = Class "PowerTime"

   Select Case udtCfg.Unit
   Case %UNIT_DATE, %UNIT_DATE_YEARMONTH, %UNIT_DATE_MONTHDAY

      Function = IsItDate(oCurrent, oValue, udtCfg)

   Case %UNIT_DAY_WEEK, %UNIT_DATE_YEAR, %UNIT_DATE_MONTH, %UNIT_DATE_DAY
   ' Special case of IsItDate, second parameter isn't a IPowerTime object

      Function = IsItDate(oCurrent, Val(sValue), udtCfg)

   Case %UNIT_TIME, %UNIT_TIME_HOURMINUTE

      Function = IsItTime(oCurrent, oValue, udtCfg)

   Case %UNIT_TIME_HOUR, %UNIT_TIME_MINUTE, %UNIT_TIME_SECOND
   ' Special case of IsItTime, second parameter isn't a IPowerTime object

      Function = IsItTime(oCurrent, Val(sValue), udtCfg)

   End Select

End Function
'------------------------------------------------------------------------------

Function IsItDate(ByVal oCurrent As IPowerTime, ByVal vntValue As Variant, _
   ByVal udtCfg As CfgTYPE) As Long
'------------------------------------------------------------------------------
'Purpose  : Evaluates the date related parameters
'
'Prereq.  : -
'Parameter: oCurrent - Current date/time
'           vntValue - Vales passed by command line (PowerTime object or number)
'           udtCfg   - Parameters passed to program
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lValue1, lValue2, lCurrent1, lCurrent2 As Long
   Local oValue As IPowerTime

   Select Case udtCfg.Unit

   Case %UNIT_DATE
   ' *** Full date

      ' Get the result
      Let oValue = vntValue

      oCurrent.DateDiff oValue, lValue1, ByVal 0&, ByVal 0&, ByVal 0&

      Select Case udtCfg.Compare
      Case < -1
      ' Before
         If lValue1 = -1 Then
            Function = %True
         Else
            Function = %False
         End If

      Case -1
      ' Before or equal
         If (lValue1 = -1) Or (lValue1 = 0) Then
            Function = %True
         Else
            Function = %False
         End If
      Case 0
      ' Equal
         If lValue1 = 0 Then
            Function = %True
         Else
            Function = %False
         End If
      Case 1
      ' After or equal
         If (lValue1 = 1) Or (lValue1 = 0) Then
            Function = %True
         Else
            Function = %False
         End If
      Case > 1
      ' After
         If lValue1 = 1 Then
            Function = %True
         Else
            Function = %False
         End If
      End Select

   Case %UNIT_DATE_YEAR, %UNIT_DATE_MONTH, %UNIT_DATE_DAY
   ' *** Year only (yyyy), month only (mm), day only (dd)

      ' Let oValue = vntValue

      lValue1 = Variant#(vntValue)
      lCurrent1 = Switch&(udtCfg.Unit = %UNIT_DATE_YEAR, oCurrent.Year, udtCfg.Unit = %UNIT_DATE_MONTH, oCurrent.Month, udtCfg.Unit = %UNIT_DATE_DAY, oCurrent.Day)

      Select Case udtCfg.Compare
      Case < -1
      ' Before
         If lValue1 < lCurrent1 Then
            Function = %True
         Else
            Function = %False
         End If

      Case -1
      ' Before or equal
         If lValue1 <= lCurrent1 Then
            Function = %True
         Else
            Function = %False
         End If
      Case 0
      ' Equal
         If lValue1 = lCurrent1 Then
            Function = %True
         Else
            Function = %False
         End If
      Case 1
      ' After or equal
         If lValue1 >= lCurrent1 Then
            Function = %True
         Else
            Function = %False
         End If
      Case > 1
      ' After
         If lValue1 > lCurrent1 Then
            Function = %True
         Else
            Function = %False
         End If
      End Select

   Case %UNIT_DATE_YEARMONTH, %UNIT_DATE_MONTHDAY
   ' *** Year and month (yyyymm) or month and day (mmdd)
   ' !!! Note !!!
   ' For this comparison, we consider the "Before" and "After" evaluation to
   ' be True, if the higher date part (year for %UNIT_DATE_YEARMONTH and
   ' month for %UNIT_DATE_MONTHDAY) is less/greater than or equal to the current date.

      Let oValue = vntValue

      If udtCfg.Unit = %UNIT_DATE_YEARMONTH Then
         lValue1 = oValue.Year : lValue2 = oValue.Month
         lCurrent1 = oCurrent.Year : lCurrent2 = oCurrent.Month
      ElseIf udtCfg.Unit = %UNIT_DATE_MONTHDAY Then
         lValue1 = oValue.Month : lValue2 = oValue.Day
         lCurrent1 = oCurrent.Month : lCurrent2 = oCurrent.Day
      End If

      Select Case udtCfg.Compare
      Case < -1
      ' Before
         If (lValue1 < lCurrent1) And (lValue2 <= lCurrent2) Then
            Function = %True
         Else
            Function = %False
         End If

      Case -1
      ' Before or equal
         If (lValue1 <= lCurrent1) And (lValue2 <= lCurrent2) Then
            Function = %True
         Else
            Function = %False
         End If
      Case 0
      ' Equal
         If (lValue1 = lCurrent1) And (lValue2 = lCurrent2) Then
            Function = %True
         Else
            Function = %False
         End If
      Case 1
      ' After or equal
         If (lValue1 >= lCurrent1) And (lValue2 >= lCurrent2) Then
            Function = %True
         Else
            Function = %False
         End If
      Case > 1
      ' After
         If (lValue1 > lCurrent1) And (lValue2 >= lCurrent2) Then
            Function = %True
         Else
            Function = %False
         End If
      End Select

   Case %UNIT_DAY_WEEK
   ' Special case: vntValue is not a IPowerTime object, but a Long (Day of week)

      lValue1 = Variant#(vntValue) + udtCfg.FDW
      lCurrent1 = oCurrent.DayOfWeek + udtCfg.FDW

      Select Case udtCfg.Compare
      Case < -1
      ' Before
         If lValue1 < lCurrent1 Then
            Function = %True
         Else
            Function = %False
         End If

      Case -1
      ' Before or equal
         If lValue1 <= lCurrent1 Then
            Function = %True
         Else
            Function = %False
         End If
      Case 0
      ' Equal
         If lValue1 = lCurrent1 Then
            Function = %True
         Else
            Function = %False
         End If
      Case 1
      ' After or equal
         If lValue1 >= lCurrent1 Then
            Function = %True
         Else
            Function = %False
         End If
      Case > 1
      ' After
         If lValue1 >= lCurrent1 Then
            Function = %True
         Else
            Function = %False
         End If
      End Select

   End Select

End Function
'------------------------------------------------------------------------------

Function IsItTime(ByVal oCurrent As IPowerTime, ByVal vntValue As Variant, _
   ByVal udtCfg As CfgTYPE) As Long
'------------------------------------------------------------------------------
'Purpose  : Evaluates the parameter /u=time
'
'Prereq.  : -
'Parameter: oCurrent - Current date/time
'           oValue   - Vales passed by command line
'           lCompare - Comparison to perform
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lSign, lDays As Long
   Local lValue1, lValue2, lCurrent1, lCurrent2 As Long
   Local oValue As IPowerTime

   ' Get the result
   Let oValue = vntValue
   oCurrent.TimeDiff oValue, lSign, lDays

   Select Case udtCfg.Unit

   Case %UNIT_TIME
   ' Full time as (hhnnss)

      Select Case udtCfg.Compare
      Case < -1
      ' Before
         If lSign = -1 Then
            Function = %True
         Else
            Function = %False
         End If

      Case -1
      ' Before or equal
         If (lSign = -1) Or (lSign = 0) Then
            Function = %True
         Else
            Function = %False
         End If
      Case 0
      ' Equal
         If lSign = 0 Then
            Function = %True
         Else
            Function = %False
         End If
      Case 1
      ' After or equal
         If (lSign = 1) Or (lSign = 0) Then
            Function = %True
         Else
            Function = %False
         End If
      Case > 1
      ' After
         If lSign = 1 Then
            Function = %True
         Else
            Function = %False
         End If
      End Select


   Case %UNIT_TIME_HOUR, %UNIT_TIME_MINUTE, %UNIT_TIME_SECOND
   ' *** Hour only (hh), minute only (nn), second only (ss)

      ' Let oValue = vntValue
      ' lValue1 = Switch&(udtCfg.Unit = %UNIT_TIME_HOUR, oValue.Hour, udtCfg.Unit = %UNIT_TIME_MINUTE, oValue.Minute, udtCfg.Unit = %UNIT_TIME_SECOND, oValue.Second)
      lValue1 = Variant#(vntValue)
      lCurrent1 = Switch&(udtCfg.Unit = %UNIT_TIME_HOUR, oCurrent.Hour, udtCfg.Unit = %UNIT_TIME_MINUTE, oCurrent.Minute, udtCfg.Unit = %UNIT_TIME_SECOND, oCurrent.Second)

      Select Case udtCfg.Compare
      Case < -1
      ' Before
         If lCurrent1 < lValue1 Then
            Function = %True
         Else
            Function = %False
         End If

      Case -1
      ' Before or equal
         If lCurrent1 <= lValue1 Then
            Function = %True
         Else
            Function = %False
         End If
      Case 0
      ' Equal
         If lCurrent1 = lValue1 Then
            Function = %True
         Else
            Function = %False
         End If
      Case 1
      ' After or equal
         If lCurrent1 >= lValue1 Then
            Function = %True
         Else
            Function = %False
         End If
      Case > 1
      ' After
         If lCurrent1 > lValue1 Then
            Function = %True
         Else
            Function = %False
         End If
      End Select


   Case %UNIT_TIME_HOURMINUTE
   ' *** Hour and minute (hhnn)
   ' !!! Note !!!
   ' For this comparison, we consider the "Before" and "After" evaluation to
   ' be True, if the hour part is less/greater than or equal to the current time.

      Let oValue = vntValue

      If udtCfg.Unit = %UNIT_TIME_HOURMINUTE Then
         lValue1 = oValue.Hour : lValue2 = oValue.Minute
         lCurrent1 = oCurrent.Hour : lCurrent2 = oCurrent.Minute
'      ElseIf udtCfg.Unit = %UNIT_MINUTE_SECOND Then
'         lValue1 = oValue.Minute : lValue2 = oValue.Second
'         lCurrent1 = oCurrent.Minute : lCurrent2 = oCurrent.Second
      End If

      Select Case udtCfg.Compare
      Case < -1
      ' Before
         If (lValue1 < lCurrent1) And (lValue2 <= lCurrent2) Then
            Function = %True
         Else
            Function = %False
         End If

      Case -1
      ' Before or equal
         If (lValue1 <= lCurrent1) And (lValue2 <= lCurrent2) Then
            Function = %True
         Else
            Function = %False
         End If
      Case 0
      ' Equal
         If (lValue1 = lCurrent1) And (lValue2 = lCurrent2) Then
            Function = %True
         Else
            Function = %False
         End If
      Case 1
      ' After or equal
         If (lValue1 >= lCurrent1) And (lValue2 >= lCurrent2) Then
            Function = %True
         Else
            Function = %False
         End If
      Case > 1
      ' After
         If (lValue1 > lCurrent1) And (lValue2 >= lCurrent2) Then
            Function = %True
         Else
            Function = %False
         End If
      End Select

   End Select

End Function
'------------------------------------------------------------------------------

Function ValidateValue(ByVal lUnit As Long, ByVal sValue As String, ByRef o As IPowerTime) As Long
'------------------------------------------------------------------------------
'Purpose  : Validates the value passed via command line in relation to the
'           unit.
'
'Prereq.  : -
'Parameter: lUnit    - Unit passed via parameter (/u)
'           sValue   - Value or this unit
'           o        - (ByRef!) IPowerTime object filled with relevant parts
'                      of sValue
'Returns  : True / False, depending on the fact if the value makes sense
'           in regards to the unit
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lTemp As Long
   Local oPTTemp As IPowerTime
   Let oPTTemp = Class "PowerTime"


   ' Assume false
   Function = %False

   Select Case lUnit

   Case %UNIT_DATE
   ' Date format: yyyymmdd

      If Len(sValue) <> 8 Then
         Exit Function
      End If

      ' All numbers?
      If Verify(sValue, "0123456789") > 0 Then
         Exit Function
      End If

      ' OK, but is it a valid date?
      Try
         oPTTemp.NewDate Val(Left$(sValue, 4)), Val(Mid$(sValue, 5, 2)), Val(Right$(sValue, 2))
      Catch
         ErrClear
         Exit Function
      End Try

      ' Looks good
      Let o = oPTTemp


   Case %UNIT_DATE_YEARMONTH
   ' Date format: yyyymm

      If Len(sValue) <> 6 Then
         Exit Function
      End If

      ' All numbers?
      If Verify(sValue, "0123456789") > 0 Then
         Exit Function
      End If

      ' OK, but is it a valid date?
      Try
         oPTTemp.NewDate Val(Left$(sValue, 4)), Val(Right$(sValue, 2)), 1
      Catch
         ErrClear
         Exit Function
      End Try

      ' Looks good
      Let o = oPTTemp


   Case %UNIT_DATE_MONTHDAY
   ' Date format: mmdd

      If Len(sValue) <> 4 Then
         Exit Function
      End If

      ' All numbers?
      If Verify(sValue, "0123456789") > 0 Then
         Exit Function
      End If

      ' OK, but is it a valid date?
      Try
         oPTTemp.NewDate 1900, Val(Left$(sValue, 2)), Val(Right$(sValue, 2))
      Catch
         ErrClear
         Exit Function
      End Try

      ' Looks good
      Let o = oPTTemp


   Case %UNIT_DATE_DAY
   ' Day as a number

      ' All numbers?
      If Verify(sValue, "0123456789") > 0 Then
         Exit Function
      End If

      ' Obviously within 1-31
      lTemp = Val(sValue)
      If lTemp < 1 Or lTemp > 31 Then
         Exit Function
      End If


   Case %UNIT_DATE_MONTH
   ' Month as a number

      ' All numbers?
      If Verify(sValue, "0123456789") > 0 Then
         Exit Function
      End If

      ' Obviously within 1-12
      lTemp = Val(sValue)
      If lTemp < 1 Or lTemp > 12 Then
         Exit Function
      End If


   Case %UNIT_DATE_YEAR
   ' Year as number, ideally as yyyy,
   ' otherwise the evaluation is always false

      ' All numbers?
      If Verify(sValue, "0123456789") > 0 Then
         Exit Function
      End If


   Case %UNIT_TIME
   ' Time format is hhmmss (24 h)

      If (Len(sValue) <> 6) Then
         Exit Function
      End If

      ' All numbers?
      If Verify(sValue, "0123456789") > 0 Then
         Exit Function
      End If

      ' OK, but is it a valid time?
      Try
         oPTTemp.NewTime Val(Left$(sValue, 2)), Val(Mid$(sValue, 5, 2)), Val(Right$(sValue, 2))
      Catch
         ErrClear
         Exit Function
      End Try

      ' Looks good
      Let o = oPTTemp


   Case %UNIT_TIME_HOUR, %UNIT_TIME_MINUTE, %UNIT_TIME_SECOND
   ' Time format: (h)h or (m)m or (s)s

      If Len(sValue) < 1 Or Len(sValue) > 2 Then
         Exit Function
      End If

      ' All numbers?
      If Verify(sValue, "0123456789") > 0 Then
         Exit Function
      End If

      If lUnit = %UNIT_TIME_HOUR Then
         ' Obviously 00-23
         If Val(sValue) < 0 Or Val(sValue) > 23 Then
            Exit Function
         End If
      ElseIf (lUnit = %UNIT_TIME_MINUTE) Or (lUnit = %UNIT_TIME_SECOND) Then
         ' Obviously 00-59
         If Val(sValue) < 0 Or Val(sValue) > 59 Then
            Exit Function
         End If
      End If


   Case %UNIT_TIME_HOURMINUTE
   ' Time format: hhmm

      If Len(sValue) < 1 Or Len(sValue) > 4 Then
         Exit Function
      End If

      ' All numbers?
      If Verify(sValue, "0123456789") > 0 Then
         Exit Function
      End If

      ' Hour, obviously 00-23
      If Val(Left$(sValue, 2)) < 0 Or Val(Left$(sValue, 2)) > 23 Then
         Exit Function
      End If

      ' Minute, obviously 00-59
      If Val(Right$(sValue, 2)) < 0 Or Val(Right$(sValue, 2)) > 59 Then
         Exit Function
      End If


   Case %UNIT_DAY_WEEK
   ' Day of week as 0-6, where 0=Sunday to 6=Saturday

      ' All numbers?
      If Verify(sValue, "0123456789") > 0 Then
         Exit Function
      End If

      ' Within 0-6, see above
      lTemp = Val(sValue)
      If lTemp < 0 Or lTemp > 0 Then
         Exit Function
      End If

   End Select


   ' If we reach this point, the validation succeeded
   Function = %True

End Function
'---------------------------------------------------------------------------

Function GetCompareStr(ByVal lValue As Long) As String
'------------------------------------------------------------------------------
'Purpose  : Returns a human readable interpretation of the comparison operator
'
'Prereq.  : -
'Parameter: lValue   - Comparison passed via parameter (/c)
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Function = Switch$(lValue < -1, "Before", lValue = -1, "Before or exact", lValue = 0, "Exact", lValue = 1, "After or exact", lValue > 1, "After")
End Function
'---------------------------------------------------------------------------

Function GetUnit(ByVal sUnit As String, ByRef lUnit As Long, Optional vntValue As Variant) As String
'------------------------------------------------------------------------------
'Purpose  : Returns a human readable interpretation of the unit
'
'Prereq.  : -
'Parameter: sUnit    - Unit passed via parameter (/u)
'           lUnit    - (ByRef!) Returns the unit also as numerical value for
'                      easier handling later on
'           vntValue - Applies only if lUnit = %UNIT_DAY_FDW.
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Select Case LCase$(sUnit)

   Case "date"
      lUnit = %UNIT_DATE
      Function = "Date"

   Case "d", "day"
      lUnit = %UNIT_DATE_DAY
      Function = "Day"

   Case "wd", "weekday"
      lUnit = %UNIT_DAY_WEEK
      Function = "Day of week"
   ' ToDo: implement week number, see http://zijlema.basicguru.eu/wrongweek.html
   'Case "wn", "weeknumber"

   Case "m", "month"
      lUnit = %UNIT_DATE_MONTH
      Function = "Month"

   Case "y", "year"
      lUnit = %UNIT_DATE_YEAR
      Function = "Year"

   Case "ym", "yearmonth"
      lUnit = %UNIT_DATE_YEARMONTH
      Function = "Year and month"

   Case "md", "monthday"
      lUnit = %UNIT_DATE_MONTHDAY
      Function = "Month and day"

   Case "fdw", "firstdayofweek"
      Local lDay As Long
      If IsMissing(vntValue) Then
         lDay = %Weekday.Monday
      ElseIf (Variant#(vntValue) < %Weekday.Sunday) Or (Variant#(vntValue) > %Weekday.Saturday) Then
         lDay = %Weekday.Monday
      Else
         lDay = Variant#(vntValue)
      End If
      lUnit = lDay
      Function = Choose$(lDay + 1, "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")

   Case "time"
      lUnit = %UNIT_TIME
      Function = "Time"

   Case "h", "hour"
      lUnit = %UNIT_TIME_HOUR
      Function = "Hour"

   Case "n", "minute"
      lUnit = %UNIT_TIME_MINUTE
      Function = "Minute"

   Case "s", "second"
      lUnit = %UNIT_TIME_SECOND
      Function = "Second"

   Case "hm", "hourminute"
      lUnit = %UNIT_TIME_HOURMINUTE
      Function = "Hour and minute"

   Case Else
      lUnit = %UNIT_OTHER
      Function = "Unknown unit"
   End Select

End Function
'---------------------------------------------------------------------------

Function GetCurrentDateTimeStr(ByVal o As IPowerTime) As String
'------------------------------------------------------------------------------
'Purpose  : Create a detailed date/time string
'
'Prereq.  : -
'Parameter: o  - Current date/time
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sResult As String

   ' Current day. We'll stick to English terms
   sResult = o.DayOfWeekString(LCIDFromLangID(%LANG_ENGLISH)) & " (" & Format$(o.DayOfWeek) & "), "
   ' Date
   sResult &= Format$(o.Year, "0000") & Format$(o.Month, "00") & Format$(o.Day, "00") & " (" & o.DateString & "), "
   ' Time
   sResult &= Format$(o.Hour, "00") & Format$(o.Minute, "00") & Format$(o.Second, "00") & " (" & o.TimeStringFull & ")"

   Function = sResult

End Function
'---------------------------------------------------------------------------

Sub ShowHelp

   ' Valid CLI parameters are:
   ' /c or /compare
   ' /u= or /unit
   ' /v= or /value
   ' /fdw or /firstdayofweek

   Con.StdOut ""
   Con.StdOut "IsIt"
   Con.StdOut "----"
   Con.StdOut "IsIt compares the current date or time using the comparison operator passed by /compare to the respective value passed by /value."
   Con.StdOut "Depending on the result of the comparison, it sets the ERRORLEVEL environment variable on exit, where"
   Con.StdOut ""
   Con.StdOut "   ERRORLEVEL =   0: the comparison failed/didn't match"
   Con.StdOut "   ERRORLEVEL =   1: the comparison succeeded/matched"
   Con.StdOut "   ERRORLEVEL = 254: invalid paramater"
   Con.StdOut "   ERRORLEVEL = 255: an error occurred during program execution"
   Con.StdOut ""
   Con.StdOut "Usage:   IsIt /compare=<comparison operator> /unit=<unit/part of date/time to compare> /value=<compare current date or time against this value>"
   Con.StdOut "            [/fdw=<first day of week>]"
   Con.StdOut ""
   Con.StdOut "e.g.     IsIt /c=-2 /u=date /v=20220321 = evaluates if the current date is before (/c=-2) 21th March 2022 (/v=20220321)."
   Con.StdOut "         IsIt /c=1 /u=time /v=233000 = evaluates if the current time is exact or after (/c=1) half to midnight (/v=233000)."
   Con.StdOut ""
   Con.StdOut "Parameters"
   Con.StdOut "----------"
   Con.StdOut "/c or /compare = Comparison operations to perform.
   Con.StdOut "                 Valid values are:"
   Con.StdOut "                 -2 - Before"
   Con.StdOut "                 -1 - Before or exact"
   Con.StdOut "                  0 - Exact"
   Con.StdOut "                  1 - After or exact"
   Con.StdOut "                  2 - After"
   Con.StdOut ""
   Con.StdOut "/u or /unit    = Unit/part of the current date or time which should be compared to the value passed."
   Con.StdOut "                 Valid values are:"
   Con.StdOut "                 Date"
   Con.StdOut "                 ----"
   Con.StdOut "                 date            - Full date (year, month and day)"
   Con.StdOut "                 y or year       - Year only"
   Con.StdOut "                 m or month      - Month only"
   Con.StdOut "                 d or day        - Day only"
   Con.StdOut "                 ym or yearmonth - Year and month"
   Con.StdOut "                 md or monthday  - Month and day"
   Con.StdOut "                 wd or weekday   - Day of week (Su-Sa)"
   Con.StdOut "                 Time"
   Con.StdOut "                 ----"
   Con.StdOut "                 time - Full time (hour, minute and second)"
   Con.StdOut "                 h or hour          - Hour only"
   Con.StdOut "                 n or minute        - Minute only"
   Con.StdOut "                 s or second        - Second only"
   Con.StdOut "                 hm or hourminute   - Hour and minute"
   ' Con.StdOut "                 ms or minutesecond - Minute and second"
   Con.StdOut ""
   Con.StdOut "/v or /value   = Value against which the current date/time should be compared."
   Con.StdOut "                 The format for passing values obviously depends on the unit passed (see above)."
   Con.StdOut "                 Valid values are:"
   Con.StdOut "                 Date"
   Con.StdOut "                 ----"
   Con.StdOut "                 for /u=date - yyyymmdd"
   Con.StdOut "                 for /u=y    - yyyy"
   Con.StdOut "                 for /u=m    - mm"
   Con.StdOut "                 for /u=d    - dd"
   Con.StdOut "                 for /u=ym   - yyyymm"
   Con.StdOut "                 for /u=md   - mmdd"
   Con.StdOut "                 for /u=wd   - 0 (Sunday) to 6 (Saturday)"
   Con.StdOut "                 Time"
   Con.StdOut "                 ----"
   Con.StdOut "                 for /u=time - hhnnss"
   Con.StdOut "                 for /u=h    - hh"
   Con.StdOut "                 for /u=n    - nn"
   Con.StdOut "                 for /u=s    - ss"
   Con.StdOut "                 for /u=hm   - hhnn"
'   Con.StdOut "                 for /u=ns   - nnss"
   Con.StdOut ""
   Con.StdOut "   Where:"
   Con.StdOut "       y - Year"
   Con.StdOut "       m - Month"
   Con.StdOut "       d - Day"
   Con.StdOut "       h - Hour (format 24 HH)"
   Con.StdOut "       n - Minute"
   Con.StdOut "       s - Second"
   Con.StdOut ""
   Con.StdOut "/fdw or /firstdayofweek = (Optional) Set which day is considered to be the first day of the week for /unit=wd."
   Con.StdOut "                          Valid values are 0 (Sunday) to 6 (Saturday). If omitted, 1 (Monday) is the default value."
   Con.StdOut ""

End Sub
'---------------------------------------------------------------------------

Function ErrString(ByVal lErr As Long, Optional ByVal vntPrefix As Variant) As String
'------------------------------------------------------------------------------
'Purpose  : Returns an formatted error string from an (PB) error number
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 12.02.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sPrefix As String

   If Not IsMissing(vntPrefix) Then
      sPrefix = Variant$(vntPrefix)
   End If

   ErrString = sPrefix & Format$(lErr) & " - " & Error$(lErr)

End Function
'------------------------------------------------------------------------------

Function LCIDFromLangID (ByVal dwLangID As Dword) As Dword

   LCIDFromLangID = MAKELCID(dwLangID, %SORT_DEFAULT)

End Function
'---------------------------------------------------------------------------

Function LocaleString Alias "LocaleString" (ByVal dwLCID As Dword, ByVal eInfo As Long) Export As String

  Dim szLocale As AsciiZ * 11
  Dim nLen As Long

  ' GetUserDefaultLCID()
  nLen = GetLocaleInfo(dwLCID, eInfo, szLocale, SizeOf(szLocale))
  LocaleString = Left$(szLocale, nLen - 1)

End Function
'---------------------------------------------------------------------------
