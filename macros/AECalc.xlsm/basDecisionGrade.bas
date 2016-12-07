Attribute VB_Name = "basDecisionGrade"
'////////////////////////////////////////////////////////////////////////////////////////
'Name         :basDecisionGrade
'Explanation  :
'Date created :2016/02/10 sakaguchi,yamamoto
'　　　　　　 :2016/12/07 sakaguchi (Add WBC_Plus_10e9L)
'////////////////////////////////////////////////////////////////////////////////////////

Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :WBC_Minus_mm3    白血球減少
'Argument     :dblResult        TestValue
'             :dblLLN           Lower Limit
'Return Value :Grade
'Date created :2016/02/09 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Function WBC_Minus_mm3(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 1000 Then
    lngGrade = 4
  ElseIf dblResult < 2000 Then
    lngGrade = 3
  ElseIf dblResult < 3000 Then
    lngGrade = 2
  ElseIf (3000 < dblLLN) And (dblResult < dblLLN) Then
    lngGrade = 1
  End If
  
  WBC_Minus_mm3 = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :WBC_Minus_10e9L  白血球減少
'Argument     :dblResult        TestValue
'             :dblLLN           Lower Limit
'Return Value :Grade
'Date created :2016/02/09 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Function WBC_Minus_10e9L(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  Select Case dblResult
    Case Is < 1
      lngGrade = 4
    Case Is < 2
      lngGrade = 3
    Case Is < 3
      lngGrade = 2
    Case Else
      If (3 < dblLLN) And (dblResult < dblLLN) Then
        lngGrade = 1
      End If
  End Select
  
  WBC_Minus_10e9L = lngGrade
  
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :WBC_Plus_mm3   白血球増加症
'Argument     :dblResult      TestValue
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function WBC_Plus_mm3(ByVal dblResult As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  Select Case dblResult
    Case Is > 100000
      lngGrade = 3
    Case Else
  End Select
  
  WBC_Plus_mm3 = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :WBC_Plus_10e9L 白血球増加症
'Argument     :dblResult      TestValue
'Return Value :Grade
'Date created :2016/12/07 坂口
'////////////////////////////////////////////////////////////////////////////////////////
Public Function WBC_Plus_10e9L(ByVal dblResult As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  Select Case dblResult
    Case Is > 100
      lngGrade = 3
    Case Else
  End Select
  
  WBC_Plus_10e9L = lngGrade
  

End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Hgb_Plus_gdL   ヘモグロビン増加
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'             :dblHgb_gdL     Base line
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Hgb_Plus_gdL(ByVal dblResult As Double, ByVal dblULN As Double, ByVal dblHgb_gdL As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblHgb_gdL >= dblULN Then
  
    Select Case dblResult - dblHgb_gdL
      Case Is > 4
        lngGrade = 3
      Case Is > 2
        lngGrade = 2
      Case Is > 0
        lngGrade = 1
      Case Else
    End Select
  
  Else
      
    Select Case dblResult - dblULN
      Case Is > 4
        lngGrade = 3
      Case Is > 2
        lngGrade = 2
      Case Is > 0
        lngGrade = 1
      Case Else
    End Select
      
      
  End If
   
  Hgb_Plus_gdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Hgb_Plus_mgL   ヘモグロビン増加
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'             :dblHgb_gdL     Base line
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Hgb_Plus_mgL(ByVal dblResult As Double, ByVal dblULN As Double, ByVal dblHgb_mgL As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblHgb_mgL >= dblULN Then
  
    Select Case dblResult - dblHgb_mgL
      Case Is > 40000
        lngGrade = 3
      Case Is > 20000
        lngGrade = 2
      Case Is > 0
        lngGrade = 1
      Case Else
    End Select
  Else
     Select Case dblResult - dblULN
      Case Is > 40000
        lngGrade = 3
      Case Is > 20000
        lngGrade = 2
      Case Is > 0
        lngGrade = 1
      Case Else
    End Select
  
  End If
   
  Hgb_Plus_mgL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Hgb_Minus_gdL    貧血
'Argument     :dblResult        TestValue
'             :dblLLN           Lower Limit
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'             :2016/12/02 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Hgb_Minus_gdL(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 8 Then '2016/12/02 sakaguchi
    lngGrade = 3
  ElseIf dblResult >= 8 And dblResult < 10 Then
    lngGrade = 2
  ElseIf dblResult >= 10 And dblResult < dblLLN Then
    lngGrade = 1
  End If

  Hgb_Minus_gdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Hgb_Minus_mgL  貧血
'Argument     :dblResult      TestValue
'             :dblLLN         Lower Limit
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'             :2016/12/02 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Hgb_Minus_mgL(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 80000 Then '2016/12/02 sakaguchi
    lngGrade = 3
  ElseIf dblResult >= 80000 And dblResult < 100000 Then
    lngGrade = 2
  ElseIf dblResult >= 100000 And dblResult < dblLLN Then
    lngGrade = 1
  End If

  Hgb_Minus_mgL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :PLT_Minus_mm3    血小板減少
'Argument     :dblResult        TestValue
'             :dblLLN           Lower Limit
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function PLT_Minus_mm3(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 25000 Then
    lngGrade = 4
  ElseIf dblResult < 50000 Then
    lngGrade = 3
  ElseIf dblResult < 75000 Then
    lngGrade = 2
  ElseIf dblResult < dblLLN Then
    lngGrade = 1
  End If

  PLT_Minus_mm3 = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :PLT_Minus_10e9L    血小板減少
'Argument     :dblResult          TestValue
'             :dblLLN             Lower Limit
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function PLT_Minus_10e9L(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 25 Then
    lngGrade = 4
  ElseIf dblResult < 50 Then
    lngGrade = 3
  ElseIf dblResult < 75 Then
    lngGrade = 2
  ElseIf dblResult < dblLLN Then
    lngGrade = 1
  End If

  PLT_Minus_10e9L = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Ne_Minus_Per1        好中球数減少
'Argument     :dblResult            TestValue
'             :dblLLN               Lower Limit
'             :dblResult_WBC_mm3    白血球数
'             :dblLLN_WBC_mm3       白血球数下限値
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Ne_Minus_Per1(ByVal dblResult As Double, ByVal dblLLN As Double, ByVal dblResult_WBC_mm3 As Double, ByVal dblLLN_WBC_mm3 As Double) As Long
  Dim lngGrade        As Long
  Dim dblResult_mm3   As Long
  Dim dblLLN_mm3   As Long

  lngGrade = 0
  dblResult_mm3 = dblResult * dblResult_WBC_mm3 * 0.01
  dblLLN_mm3 = dblLLN * dblLLN_WBC_mm3 * 0.01
  
  If dblResult_mm3 < 500 Then
    lngGrade = 4
  ElseIf dblResult_mm3 < 1000 Then
    lngGrade = 3
  ElseIf dblResult_mm3 < 1500 Then
    lngGrade = 2
  ElseIf dblResult_mm3 < dblLLN_mm3 Then
    lngGrade = 1
  End If

  Ne_Minus_Per1 = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Ne_Minus_Per2        好中球数減少
'Argument     :dblResult            TestValue
'             :dblLLN               Lower Limit
'             :dblResult_WBC_10e9L  白血球数
'             :dblLLN_WBC_10e9L     白血球数下限値
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Ne_Minus_Per2(ByVal dblResult As Double, ByVal dblLLN As Double, ByVal dblResult_WBC_10e9L As Double, ByVal dblLLN_WBC_10e9L As Double) As Long
  Dim lngGrade        As Long
  Dim dblResult_10e9L   As Long
  Dim dblLLN_10e9L   As Long

  lngGrade = 0
  dblResult_10e9L = dblResult * dblResult_WBC_10e9L * 0.01
  dblLLN_10e9L = dblLLN * dblLLN_WBC_10e9L * 0.01
  
  If dblResult_10e9L < 0.5 Then
    lngGrade = 4
  ElseIf dblResult_10e9L < 1 Then
    lngGrade = 3
  ElseIf dblResult_10e9L < 1.5 Then
    lngGrade = 2
  ElseIf dblResult_10e9L < dblLLN_10e9L Then
    lngGrade = 1
  End If

  Ne_Minus_Per2 = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Ly_Minus_Per1        リンパ球数減少
'Argument     :dblResult            TestValue
'             :dblLLN               Lower Limit
'             :dblResult_WBC_mm3    白血球数
'             :dblLLN_WBC_mm3       白血球数下限値
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Ly_Minus_Per1(ByVal dblResult As Double, ByVal dblLLN As Double, ByVal dblResult_WBC_mm3 As Double, ByVal dblLLN_WBC_mm3 As Double) As Long
  Dim lngGrade        As Long
  Dim dblResult_mm3   As Long
  Dim dblLLN_mm3   As Long

  lngGrade = 0
  dblResult_mm3 = dblResult * dblResult_WBC_mm3 * 0.01
  dblLLN_mm3 = dblLLN * dblLLN_WBC_mm3 * 0.01
  
  If dblResult_mm3 < 200 Then
    lngGrade = 4
  ElseIf dblResult_mm3 < 500 Then
    lngGrade = 3
  ElseIf dblResult_mm3 < 800 Then
    lngGrade = 2
  ElseIf dblResult_mm3 < dblLLN_mm3 Then
    lngGrade = 1
  End If

  Ly_Minus_Per1 = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Ly_Minus_Per2          リンパ球数減少
'Argument     :dblResult              TestValue
'             :dblLLN                 Lower Limit
'             :dblResult_WBC_10e9L    白血球数
'             :dblLLN_WBC_10e9L       白血球数下限値
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Ly_Minus_Per2(ByVal dblResult As Double, ByVal dblLLN As Double, ByVal dblResult_WBC_10e9L As Double, ByVal dblLLN_WBC_10e9L As Double) As Long
  Dim lngGrade        As Long
  Dim dblResult_10e9L   As Long
  Dim dblLLN_10e9L   As Long

  lngGrade = 0
  dblResult_10e9L = dblResult * dblResult_WBC_10e9L * 0.01
  dblLLN_10e9L = dblLLN * dblLLN_WBC_10e9L * 0.01
  
  If dblResult_10e9L < 0.2 Then
    lngGrade = 4
  ElseIf dblResult_10e9L < 0.5 Then
    lngGrade = 3
  ElseIf dblResult_10e9L < 0.8 Then
    lngGrade = 2
  ElseIf dblResult_10e9L < dblLLN_10e9L Then
    lngGrade = 1
  End If

  Ly_Minus_Per2 = lngGrade
  
End Function

'///////////////////////////////////////////////////////////////////////////////////////
'Name         :Ly_Plus_Per1         リンパ球数増加
'Argument     :dblResult            TestValue
'             :dblResult_WBC_mm3    白血球数
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Ly_Plus_Per1(ByVal dblResult As Double, ByVal dblResult_WBC_mm3 As Double) As Long
  Dim lngGrade        As Long
  Dim dblResult_mm3   As Long
 
  lngGrade = 0
  dblResult_mm3 = dblResult * dblResult_WBC_mm3 * 0.01
   
  If dblResult_mm3 > 20000 Then
    lngGrade = 3
  ElseIf dblResult_mm3 > 4000 Then
    lngGrade = 2
  End If

  Ly_Plus_Per1 = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Ly_Plus_Per2           リンパ球数増加
'Argument     :dblResult              TestValue
'             :dblResult_WBC_10e9L    白血球数
'Return Value :Grade
'Date created :2016/02/09 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Ly_Plus_Per2(ByVal dblResult As Double, ByVal dblResult_WBC_10e9L As Double) As Long
  Dim lngGrade        As Long
  Dim dblResult_10e9L   As Long
 
  lngGrade = 0
  dblResult_10e9L = dblResult * dblResult_WBC_10e9L * 0.01
   
  If dblResult_10e9L > 20 Then
    lngGrade = 3
  ElseIf dblResult_10e9L > 4 Then
    lngGrade = 2
  End If

  Ly_Plus_Per2 = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :PT_Plus_INR    INR増加
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function PT_Plus_INR(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
   
  Select Case dblResult
    Case Is > 2.5 * dblULN
      lngGrade = 3
    Case Is > 1.5 * dblULN
      lngGrade = 2
    Case Is > 1 * dblULN
      lngGrade = 1
    Case Else
  End Select
    
    
  PT_Plus_INR = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :APTT_Plus_SEC  活性化部分トロンボプラスチン時間延長
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function APTT_Plus_SEC(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
   
  Select Case dblResult
    Case Is > 2.5 * dblULN
      lngGrade = 3
    Case Is > 1.5 * dblULN
      lngGrade = 2
    Case Is > 1 * dblULN
      lngGrade = 1
    Case Else
  End Select
    
    
  APTT_Plus_SEC = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Fib_Minus_mgdL fibrinogen減少
'Argument     :dblResult      TestValue
'             :dblLLN         Lower Limit
'             :dblFib_mgdL    Base line
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Fib_Minus_mgdL(ByVal dblResult As Double, ByVal dblLLN As Double, ByVal dblFib_mgdL As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  
  If (dblResult < 0.25 * dblLLN) Or (dblResult <= dblFib_mgdL - (0.75 * dblFib_mgdL)) Or (dblResult < 50) Then
    lngGrade = 4
  ElseIf (dblResult < 0.5 * dblLLN) Or (dblResult <= dblFib_mgdL - (0.5 * dblFib_mgdL)) Then
    lngGrade = 3
  ElseIf (dblResult < 0.75 * dblLLN) Or (dblResult <= dblFib_mgdL - (0.25 * dblFib_mgdL)) Then
    lngGrade = 2
  ElseIf (dblResult < 1 * dblLLN) Or (dblResult < dblFib_mgdL) Then
    lngGrade = 1
  End If
   
  Fib_Minus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :ALB_Minus_gdL  低アルブミン血症
'Argument     :dblResult      TestValue
'             :dblLLN         Lower Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function ALB_Minus_gdL(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 2 Then
    lngGrade = 3
  ElseIf dblResult < 3 Then
    lngGrade = 2
  ElseIf dblResult < dblLLN Then
    lngGrade = 1
  End If

  ALB_Minus_gdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :ALB_Minus_gL   低アルブミン血症
'Argument     :dblResult      TestValue
'             :dblLLN         Lower Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function ALB_Minus_gL(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 20 Then
    lngGrade = 3
  ElseIf dblResult < 30 Then
    lngGrade = 2
  ElseIf dblResult < dblLLN Then
    lngGrade = 1
  End If

  ALB_Minus_gL = lngGrade
  
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Cre_Plus_mgdL    クレアチニン増加
'Argument     :dblResult        TestValue
'             :dblLLN           Upper  Limit
'             :dblCre_mgdL      Base line
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Cre_Plus_mgdL(ByVal dblResult As Double, ByVal dblULN As Double, ByVal dblCre_mgdL As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If (dblResult > 6 * dblULN) Then
    lngGrade = 4
  ElseIf (dblResult > 3 * dblULN) Or (dblResult > 3 * dblCre_mgdL) Then
    lngGrade = 3
  ElseIf (dblResult > 1.5 * dblULN) Or (dblResult > 1.5 * dblCre_mgdL) Then
    lngGrade = 2
  ElseIf (dblResult > dblULN) Or (dblResult > dblCre_mgdL) Then
    lngGrade = 1
  End If
   
  Cre_Plus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Cre_Plus2_mgdL   急性腎不全
'Argument     :dblResult        TestValue
'             :dblLLN           Upper  Limit
'             :dblCre_mgdL      Base line
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Cre_Plus2_mgdL(ByVal dblResult As Double, ByVal dblULN As Double, ByVal dblCre_mgdL As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If (dblResult > 4) Or (dblResult > 3 * dblCre_mgdL) Then
    lngGrade = 3
  ElseIf dblResult > 2 * dblCre_mgdL Then
    lngGrade = 2
  ElseIf (dblResult - dblCre_mgdL > 0.3) Or (dblResult > 1.5 * dblCre_mgdL) Then
    lngGrade = 1
  End If
   
  Cre_Plus2_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :UA_Plus_mgdL   高尿酸血症
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function UA_Plus_mgdL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > 10 Then
    lngGrade = 4
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  UA_Plus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :CHO_Plus_mgdL  コレステロール増加
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function CHO_Plus_mgdL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > 500 Then
    lngGrade = 4
  ElseIf dblResult > 400 Then
    lngGrade = 3
  ElseIf dblResult > 300 Then
    lngGrade = 2
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  CHO_Plus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Tbil_Plus_mgdL   血中ビリルビン増加
'Argument     :dblResult        TestValue
'             :dblLLN           Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Tbil_Plus_mgdL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > dblULN * 10 Then
    lngGrade = 4
  ElseIf dblResult > dblULN * 3 Then
    lngGrade = 3
  ElseIf dblResult > dblULN * 1.5 Then
    lngGrade = 2
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  Tbil_Plus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :ALP_Plus_UL    アルカリホスファターゼ増加
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function ALP_Plus_UL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > dblULN * 20 Then
    lngGrade = 4
  ElseIf dblResult > dblULN * 5 Then
    lngGrade = 3
  ElseIf dblResult > dblULN * 2.5 Then
    lngGrade = 2
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  ALP_Plus_UL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :CPK_Plus_UL    CPK増加
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function CPK_Plus_UL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > dblULN * 10 Then
    lngGrade = 4
  ElseIf dblResult > dblULN * 5 Then
    lngGrade = 3
  ElseIf dblResult > dblULN * 2.5 Then
    lngGrade = 2
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  CPK_Plus_UL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :AST_Plus_UL    アスパラギン酸アミノトランスフェラーゼ増加
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function AST_Plus_UL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > dblULN * 20 Then
    lngGrade = 4
  ElseIf dblResult > dblULN * 5 Then
    lngGrade = 3
  ElseIf dblResult > dblULN * 3 Then
    lngGrade = 2
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  AST_Plus_UL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :ALT_Plus_UL    アラニン・アミノトランスフェラーゼ増加
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function ALT_Plus_UL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > dblULN * 20 Then
    lngGrade = 4
  ElseIf dblResult > dblULN * 5 Then
    lngGrade = 3
  ElseIf dblResult > dblULN * 3 Then
    lngGrade = 2
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  ALT_Plus_UL = lngGrade
  
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :GTP_Plus_UL    GGT増加
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/12 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function GTP_Plus_UL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > dblULN * 20 Then
    lngGrade = 4
  ElseIf dblResult > dblULN * 5 Then
    lngGrade = 3
  ElseIf dblResult > dblULN * 2.5 Then
    lngGrade = 2
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  GTP_Plus_UL = lngGrade
  
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Na_Plus_mEqL  高ナトリウム血症
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Na_Plus_mEqL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > 160 Then
    lngGrade = 4
  ElseIf dblResult > 155 Then
    lngGrade = 3
  ElseIf dblResult > 150 Then
    lngGrade = 2
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  Na_Plus_mEqL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Na_Minus_mEqL  低ナトリウム血症
'Argument     :dblResult      TestValue
'             :dblLLN         Lower Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Na_Minus_mEqL(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 120 Then
    lngGrade = 4
  ElseIf dblResult < 130 Then
    lngGrade = 3
  ElseIf dblResult < dblLLN Then
    lngGrade = 1
  End If
   
  Na_Minus_mEqL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :K_Plus_mEqL  高カリウム血症
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function K_Plus_mEqL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > 7 Then
    lngGrade = 4
  ElseIf dblResult > 6 Then
    lngGrade = 3
  ElseIf dblResult > 5.5 Then
    lngGrade = 2
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  K_Plus_mEqL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :K_Minus_mEqL   低カリウム血症
'Argument     :dblResult      TestValue
'             :dblLLN         Lower Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function K_Minus_mEqL(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 2.5 Then
    lngGrade = 4
  ElseIf dblResult < 3 Then
    lngGrade = 3
  ElseIf dblResult < dblLLN Then
    lngGrade = 1
  End If
   
  K_Minus_mEqL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Ca_Plus_mgdL   高カルシウム血症
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Ca_Plus_mgdL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > 13.5 Then
    lngGrade = 4
  ElseIf dblResult > 12.5 Then
    lngGrade = 3
  ElseIf dblResult > 11.5 Then
    lngGrade = 2
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  Ca_Plus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Ca_Minus_mgdL  低カルシウム血症
'Argument     :dblResult      TestValue
'             :dblLLN         Lower Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Ca_Minus_mgdL(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 6 Then
    lngGrade = 4
  ElseIf dblResult < 7 Then
    lngGrade = 3
  ElseIf dblResult < 8 Then
    lngGrade = 2
  ElseIf dblResult < dblLLN Then
    lngGrade = 1
  End If
   
  Ca_Minus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :IP_Minus_mgdL  低リン血症
'Argument     :dblResult      TestValue
'             :dblLLN         Lower Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function IP_Minus_mgdL(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 1 Then
    lngGrade = 4
  ElseIf dblResult < 2 Then
    lngGrade = 3
  ElseIf dblResult < 2.5 Then
    lngGrade = 2
  ElseIf dblResult < dblLLN Then
    lngGrade = 1
  End If
   
  IP_Minus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Mg_Plus_mgdL   高マグネシウム血症
'Argument     :dblResult      TestValue
'             :dblLLN         Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Mg_Plus_mgdL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > 8 Then
    lngGrade = 4
  ElseIf dblResult > 3 Then
    lngGrade = 3
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  Mg_Plus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Mg_Minus_mgdL  低マグネシウム血症
'Argument     :dblResult      TestValue
'             :dblLLN         Lower Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Mg_Minus_mgdL(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 0.7 Then
    lngGrade = 4
  ElseIf dblResult < 0.9 Then
    lngGrade = 3
  ElseIf dblResult < 1.2 Then
    lngGrade = 2
  ElseIf dblResult < dblLLN Then
    lngGrade = 1
  End If
   
  Mg_Minus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Gluc_Plus_mgdL   高血糖
'Argument     :dblResult        TestValue
'             :dblLLN           Upper  Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Gluc_Plus_mgdL(ByVal dblResult As Double, ByVal dblULN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult > 500 Then
    lngGrade = 4
  ElseIf dblResult > 250 Then
    lngGrade = 3
  ElseIf dblResult > 160 Then
    lngGrade = 2
  ElseIf dblResult > dblULN Then
    lngGrade = 1
  End If
   
  Gluc_Plus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :Gluc_Minus_mgdL  低血糖症
'Argument     :dblResult        TestValue
'             :dblLLN           Lower Limit
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function Gluc_Minus_mgdL(ByVal dblResult As Double, ByVal dblLLN As Double) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If dblResult < 30 Then
    lngGrade = 4
  ElseIf dblResult < 40 Then
    lngGrade = 3
  ElseIf dblResult < 55 Then
    lngGrade = 2
  ElseIf dblResult < dblLLN Then
    lngGrade = 1
  End If
   
  Gluc_Minus_mgdL = lngGrade
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :UPro_Plus      尿蛋白
'Argument     :dblResult      TestValue
'Return Value :Grade
'Date created :2016/02/10 yamamoto
'////////////////////////////////////////////////////////////////////////////////////////
Public Function UPro_Plus(ByVal strResult As String) As Long
  Dim lngGrade        As Long

  lngGrade = 0
  
  If strResult = "2+" Then
    lngGrade = 2
  ElseIf strResult = "1+" Then
    lngGrade = 1
  End If
   
  UPro_Plus = lngGrade
  
End Function


