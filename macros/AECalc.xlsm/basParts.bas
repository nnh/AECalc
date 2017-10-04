Attribute VB_Name = "basParts"
'////////////////////////////////////////////////////////////////////////////////////////
'Name         :basParts
'Explanation  :
'Date created : 2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////

Option Explicit


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :CalcAge
'Argument     :ByRef lngY　       Age
'             :ByRef lngM　       MonthOld
'             :ByVal dtBirthday   Birthday
'             :ByVal dtTestday    TestDay
'Return Value :0(Error then ErrorNumber)
'Date created :2016/02/08 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Function CalcAge(ByRef lngY As Long, ByRef lngM As Long, ByVal dtBirthday As Date, ByVal dtTestday As Date) As Long
  Dim lngMwork As Long
  
  On Error GoTo lineErr:

  CalcAge = 0
  lngY = 0
  lngM = 0
  
  '/// CalcAge
  lngY = DateDiff("yyyy", dtBirthday, dtTestday)
  If Format(dtTestday, "mmdd") < Format(dtBirthday, "mmdd") Then lngY = lngY - 1
  
  '/// CalcMonthOld
  lngMwork = DateDiff("m", dtBirthday, dtTestday)
  lngMwork = lngMwork Mod 12
  If Format(dtTestday, "dd") < Format(dtBirthday, "dd") Then
    If 0 < lngMwork Then lngMwork = lngMwork - 1 Else lngMwork = 11  '/ -1 Month  →  11 Month
  End If
  lngM = lngMwork
    
  Exit Function
lineErr:
  CalcAge = Err.Number
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :GetLngItem
'Argument     :ByRef colCollection     LongClassItemCollection
'             :ByVal strKey            UniqueKey
'Return Value :Item Value
'Date created :2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Function GetLngItem(ByRef colCollection As Collection, ByVal strKey As String) As Long
  GetLngItem = -1
  On Error Resume Next
  GetLngItem = colCollection.Item(strKey)
  On Error GoTo 0
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :AddLngItem
'Argument     :ByRef colCollection
'             :ByVal lngItem           LongClassItem
'             :ByVal strKey            UniqueKey
'Return Value :None
'Date created :2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Sub AddLngItem(ByRef colCollection As Collection, ByVal lngItem As Long, ByVal strKey)
  On Error Resume Next
  If colCollection Is Nothing Then Set colCollection = New Collection
  colCollection.Add lngItem, strKey
  On Error GoTo 0
End Sub


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :SetIsNumeric
'Argument     :ByVal strValue
'Return Value :Numeric Value
'Date created :2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Function SetIsNumeric(ByVal strValue As String) As Double
  If IsNumeric(strValue) Then
    SetIsNumeric = strValue
  Else
    SetIsNumeric = 0
  End If
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Sheet_ApplicationOff
'Argument     :strSheet
'Return Value :None
'Date created :2016/04/19 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Sub Sheet_ApplicationOff(ByVal strSheet As String)

  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual
  Worksheets(strSheet).Unprotect

End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Sheet_ApplicationOn
'Argument     :strSheet
'Return Value :None
'Date created :2016/04/19 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Sub Sheet_ApplicationOn(ByVal strSheet As String)
  
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Application.Calculation = xlCalculationAutomatic
  Worksheets(strSheet).Protect
  
End Sub
