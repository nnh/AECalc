Attribute VB_Name = "basParts"

'////////////////////////////////////////////////////////////////////////////////////////
'名　前：basParts
'説　明：
'作成日：2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////

Option Explicit


'////////////////////////////////////////////////////////////////////////////////////////
'名　前：CalcAge
'引　数：ByRef lngY　年齢
'　　　：ByRef lngM　月齢
'　　　：ByVal dtBirthday 誕生日
'　　　：ByVal dtKensaday 検査日
'戻り値：通常0(エラー時エラー番号)　年齢、月齢
'作成日：2016/02/08 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Function CalcAge(ByRef lngY As Long, ByRef lngM As Long, ByVal dtBirthday As Date, ByVal dtKensaday As Date) As Long
  Dim lngMwork As Long
  
  On Error GoTo lineErr:

  CalcAge = 0
  lngY = 0
  lngM = 0
  
  '/// 年齢計算
  lngY = DateDiff("yyyy", dtBirthday, dtKensaday)
  If Format(dtKensaday, "mmdd") < Format(dtBirthday, "mmdd") Then lngY = lngY - 1
  
  '/// 月齢計算
  lngMwork = DateDiff("m", dtBirthday, dtKensaday)
  lngMwork = lngMwork Mod 12
  If Format(dtKensaday, "dd") < Format(dtBirthday, "dd") Then
    If 0 < lngMwork Then lngMwork = lngMwork - 1 Else lngMwork = 11  '/ -1ヶ月でなく11ヶ月に
  End If
  lngM = lngMwork
    
  Exit Function
lineErr:
  CalcAge = Err.Number
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'名　前：GetLngItem
'引　数：ByRef colCollection 対象Long型コレクション
'　　　：ByVal strKey        ユニークキー
'戻り値：Ksyに対するItemがなければ-1 あればItemの数値
'作成日：2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Function GetLngItem(ByRef colCollection As collection, ByVal strKey As String) As Long
  GetLngItem = -1
  On Error Resume Next
  GetLngItem = colCollection.Item(strKey)
  On Error GoTo 0
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'名　前：AddLngItem
'引　数：ByRef colCollection 対象Long型コレクション
'　　　：ByVal lngItem       Long型Item
'　　　：ByVal strKey        ユニークキー
'戻り値：なし
'作成日：2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Sub AddLngItem(ByRef colCollection As collection, ByVal lngItem As Long, ByVal strKey)
  On Error Resume Next
  If colCollection Is Nothing Then Set colCollection = New collection
  colCollection.Add lngItem, strKey
  On Error GoTo 0
End Sub


'////////////////////////////////////////////////////////////////////////////////////////
'名　前：SetIsNumeric
'引　数：ByVal strValue
'戻り値：数値文字列なら数値　そうでないなら0
'作成日：2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Function SetIsNumeric(ByVal strValue As String) As Double
  If IsNumeric(strValue) Then
    SetIsNumeric = strValue
  Else
    SetIsNumeric = 0
  End If
End Function

Private Sub test()
  Dim y As Long
  Dim m As Long
  Call CalcAge(y, m, "1978/2/16", "2016/2/15")
  Debug.Print y & " " & m
  
  Call CalcAge(y, m, "2014/2/15", "2016/2/15")
  Debug.Print y & " " & m
  
  Call CalcAge(y, m, "2016/11/15", "2016/12/15")
  Debug.Print y & " " & m
End Sub
