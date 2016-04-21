Attribute VB_Name = "basLaboMain"
'////////////////////////////////////////////////////////////////////////////////////////
'Name         :basLaboMain
'Explanation  :
'Date created : 2016/02/10 sakaguchi
'             : 2016/04/21 sakaguchi (Add SettingRefOver20)
'////////////////////////////////////////////////////////////////////////////////////////

Option Explicit

'// Const
Private Const mcstrOver20         As String = "over20"
Private Const mclngAdult          As Long = 999
Private Const mclngLaboSttRow     As Long = 3     '/ LaboSheet  StartRow
Private Const mclngDemogSttRow    As Long = 2     '/ DemogSheet StartRow
Private Const mclngRefSttRow      As Long = 3     '/ Ref Sheet  StartRow

'// Labo Row
Private Const mcCaseNo      As Long = 1
Private Const mcTestDay     As Long = 2
Private Const mcWBC1        As Long = 3
Private Const mcWBC2        As Long = 6
Private Const mcHgb1        As Long = 9
Private Const mcHgb2        As Long = 12
Private Const mcPLT1        As Long = 15
Private Const mcPLT2        As Long = 18
Private Const mcNe          As Long = 21
Private Const mcLy          As Long = 24
Private Const mcPT          As Long = 27
Private Const mcAPTT        As Long = 30
Private Const mcFib         As Long = 33
Private Const mcALB1        As Long = 36
Private Const mcALB2        As Long = 39
Private Const mcCre         As Long = 42
Private Const mcUA          As Long = 45
Private Const mcCHO         As Long = 48
Private Const mcTbil        As Long = 51
Private Const mcALP         As Long = 54
Private Const mcCPK         As Long = 57
Private Const mcAST         As Long = 60
Private Const mcALT         As Long = 63
Private Const mcGTP         As Long = 66
Private Const mcNa          As Long = 69
Private Const mcK           As Long = 72
Private Const mcCa          As Long = 75
Private Const mcIP          As Long = 78
Private Const mcMg          As Long = 81
Private Const mcGluc        As Long = 84
Private Const mcUPro        As Long = 87

'// Ref  Row
Private Const mcLnWBC1      As Long = 4
Private Const mcLnWBC2      As Long = 6
Private Const mcLnHgb1      As Long = 8
Private Const mcLnHgb2      As Long = 10
Private Const mcLnPLT1      As Long = 12
Private Const mcLnPLT2      As Long = 14
Private Const mcLnNe        As Long = 16
Private Const mcLnLy        As Long = 18
Private Const mcLnPT        As Long = 20
Private Const mcLnAPTT      As Long = 22
Private Const mcLnFib       As Long = 24
Private Const mcLnALB1      As Long = 26
Private Const mcLnALB2      As Long = 28
Private Const mcLnCre       As Long = 30
Private Const mcLnUA        As Long = 32
Private Const mcLnCHO       As Long = 34
Private Const mcLnTbil      As Long = 36
Private Const mcLnALP       As Long = 38
Private Const mcLnCPK       As Long = 40
Private Const mcLnAST       As Long = 42
Private Const mcLnALT       As Long = 44
Private Const mcLnGTP       As Long = 46
Private Const mcLnNa        As Long = 48
Private Const mcLnK         As Long = 50
Private Const mcLnCa        As Long = 52
Private Const mcLnIP        As Long = 54
Private Const mcLnMg        As Long = 56
Private Const mcLnGluc      As Long = 58
Private Const mcLnUPro      As Long = 60

'// variable
Private mcolAgeKaisou       As Collection '/ AgeSexCollection
Private mlngMaxRow          As Long       '/ Labo MaxRows
Private mlngMaxRowDemog     As Long       '/ Demog MaxRows

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :CalcGradeMain
'Argument     :None
'Return Value :None
'Date created :2016/02/23 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub CalcGradeMain()
  Dim strCaseNum        As String
  Dim dtTestday         As Date
  Dim i                 As Long
  Dim clPatient         As clsPatient
  Dim dblTestValue      As Double
  Dim dblLLN            As Double
  Dim dblULN            As Double
  Dim dblTestValueWBC   As Double
  Dim dblLLNWBC         As Double
  Dim dblULNWBC         As Double
  Dim strTestValue      As String
  
  Set mcolAgeKaisou = GetKaisou()
  If mcolAgeKaisou Is Nothing Then Exit Sub
  
  With Worksheets("Labo")
    For i = mclngLaboSttRow To mlngMaxRow
      strCaseNum = Worksheets("Labo").Cells(i, mcCaseNo).Value
        
      dtTestday = Worksheets("Labo").Cells(i, mcTestDay).Value
    
      Set clPatient = GetPatient(strCaseNum, dtTestday)
    
      If IsReady(mcLnWBC1, mcWBC1, clPatient, i, dblLLN, dblULN) Then '/ WBC(/mm3)
        dblTestValue = .Cells(i, mcWBC1).Value
       .Cells(i, mcWBC1 + 1).Value = WBC_Plus_mm3(dblTestValue)
       .Cells(i, mcWBC1 + 2).Value = WBC_Minus_mm3(dblTestValue, dblLLN)
      
       dblTestValueWBC = dblTestValue
       dblLLNWBC = dblLLN
       If IsReady(mcLnNe, mcNe, clPatient, i, dblLLN, dblULN) Then    '/ Ne(%)
        dblTestValue = .Cells(i, mcNe).Value
        .Cells(i, mcNe + 2).Value = Ne_Minus_Per1(dblTestValue, dblLLN, dblTestValueWBC, dblLLNWBC)
       End If
       
       If IsReady(mcLnLy, mcLy, clPatient, i, dblLLN, dblULN) Then    '/ Ly(%)
        dblTestValue = .Cells(i, mcLy).Value
        .Cells(i, mcLy + 1).Value = Ly_Plus_Per1(dblTestValue, dblTestValueWBC)
        .Cells(i, mcLy + 2).Value = Ly_Minus_Per1(dblTestValue, dblLLN, dblTestValueWBC, dblLLNWBC)
       End If
       
      End If
    
      If IsReady(mcLnWBC2, mcWBC2, clPatient, i, dblLLN, dblULN) Then  '/ WBC(10e9/L)
        dblTestValue = .Cells(i, mcWBC2).Value
       .Cells(i, mcWBC2 + 2).Value = WBC_Minus_10e9L(dblTestValue, dblLLN)
       
       dblTestValueWBC = dblTestValue
       dblLLNWBC = dblLLN
       If IsReady(mcLnNe, mcNe, clPatient, i, dblLLN, dblULN) Then      '/ Ne(%)
        dblTestValue = .Cells(i, mcNe).Value
        .Cells(i, mcNe + 2).Value = Ne_Minus_Per2(dblTestValue, dblLLN, dblTestValueWBC, dblLLNWBC)
       End If
       
       If IsReady(mcLnLy, mcLy, clPatient, i, dblLLN, dblULN) Then      '/ Ly%)
        dblTestValue = .Cells(i, mcLy).Value
        .Cells(i, mcLy + 1).Value = Ly_Plus_Per2(dblTestValue, dblTestValueWBC)
        .Cells(i, mcLy + 2).Value = Ly_Minus_Per2(dblTestValue, dblLLN, dblTestValueWBC, dblLLNWBC)
       End If
      
      End If
    
      If IsReady(mcLnHgb1, mcHgb1, clPatient, i, dblLLN, dblULN) Then '/ Hgb(g/dL)
        dblTestValue = .Cells(i, mcHgb1).Value
       .Cells(i, mcHgb1 + 1).Value = Hgb_Plus_gdL(dblTestValue, dblULN, clPatient.Hgb_gdL)
       .Cells(i, mcHgb1 + 2).Value = Hgb_Minus_gdL(dblTestValue, dblLLN)
      End If
        
      If IsReady(mcLnHgb2, mcHgb2, clPatient, i, dblLLN, dblULN) Then '/ Hgb(mg/L)
        dblTestValue = .Cells(i, mcHgb2).Value
       .Cells(i, mcHgb2 + 1).Value = Hgb_Plus_mgL(dblTestValue, dblULN, clPatient.Hgb_mgL)
       .Cells(i, mcHgb2 + 2).Value = Hgb_Minus_mgL(dblTestValue, dblLLN)
      End If

      If IsReady(mcLnPLT1, mcPLT1, clPatient, i, dblLLN, dblULN) Then '/ PLT(/mm3)
        dblTestValue = .Cells(i, mcPLT1).Value
       .Cells(i, mcPLT1 + 2).Value = PLT_Minus_mm3(dblTestValue, dblLLN)
      End If
    
      If IsReady(mcLnPLT2, mcPLT2, clPatient, i, dblLLN, dblULN) Then '/ PLT(10e9/L)
        dblTestValue = .Cells(i, mcPLT2).Value
       .Cells(i, mcPLT2 + 2).Value = PLT_Minus_10e9L(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnPT, mcPT, clPatient, i, dblLLN, dblULN) Then '/ PT(PT-INR)
        dblTestValue = .Cells(i, mcPT).Value
       .Cells(i, mcPT + 1).Value = PT_Plus_INR(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnAPTT, mcAPTT, clPatient, i, dblLLN, dblULN) Then '/ APTT(sec)
        dblTestValue = .Cells(i, mcAPTT).Value
       .Cells(i, mcAPTT + 1).Value = APTT_Plus_SEC(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnFib, mcFib, clPatient, i, dblLLN, dblULN) Then '/ fib
        dblTestValue = .Cells(i, mcFib).Value
       .Cells(i, mcFib + 2).Value = Fib_Minus_mgdL(dblTestValue, dblLLN, clPatient.Fib)
      End If
      
      If IsReady(mcLnALB1, mcALB1, clPatient, i, dblLLN, dblULN) Then '/ ALB(g/dL)
        dblTestValue = .Cells(i, mcALB1).Value
       .Cells(i, mcALB1 + 2).Value = ALB_Minus_gdL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnALB2, mcALB2, clPatient, i, dblLLN, dblULN) Then '/ ALB(g/L)
        dblTestValue = .Cells(i, mcALB2).Value
       .Cells(i, mcALB2 + 2).Value = ALB_Minus_gL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnCre, mcCre, clPatient, i, dblLLN, dblULN) Then '/ Cre(mg/dL)
        dblTestValue = .Cells(i, mcCre).Value
       .Cells(i, mcCre + 1).Value = Cre_Plus_mgdL(dblTestValue, dblULN, clPatient.Cre)
       .Cells(i, mcCre + 2).Value = Cre_Plus2_mgdL(dblTestValue, dblULN, clPatient.Cre)
      End If
      
      If IsReady(mcLnUA, mcUA, clPatient, i, dblLLN, dblULN) Then     '/ UA(mg/dL)
        dblTestValue = .Cells(i, mcUA).Value
       .Cells(i, mcUA + 1).Value = UA_Plus_mgdL(dblTestValue, dblULN)
      End If
     
      If IsReady(mcLnCHO, mcCHO, clPatient, i, dblLLN, dblULN) Then   '/ T-CHO(mg/dL)
        dblTestValue = .Cells(i, mcCHO).Value
       .Cells(i, mcCHO + 1).Value = CHO_Plus_mgdL(dblTestValue, dblULN)
      End If
     
      If IsReady(mcLnTbil, mcTbil, clPatient, i, dblLLN, dblULN) Then '/ T-Tbil(mg/dL)
        dblTestValue = .Cells(i, mcTbil).Value
       .Cells(i, mcTbil + 1).Value = Tbil_Plus_mgdL(dblTestValue, dblULN)
      End If
     
      If IsReady(mcLnALP, mcALP, clPatient, i, dblLLN, dblULN) Then '/ ALT(U/L)
        dblTestValue = .Cells(i, mcALP).Value
       .Cells(i, mcALP + 1).Value = ALP_Plus_UL(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnCPK, mcCPK, clPatient, i, dblLLN, dblULN) Then '/ CPK(U/L)
        dblTestValue = .Cells(i, mcCPK).Value
       .Cells(i, mcCPK + 1).Value = CPK_Plus_UL(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnAST, mcAST, clPatient, i, dblLLN, dblULN) Then '/ AST(U/L)
        dblTestValue = .Cells(i, mcAST).Value
       .Cells(i, mcAST + 1).Value = AST_Plus_UL(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnALT, mcALT, clPatient, i, dblLLN, dblULN) Then '/ ALT(U/L)
        dblTestValue = .Cells(i, mcALT).Value
       .Cells(i, mcALT + 1).Value = ALT_Plus_UL(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnGTP, mcGTP, clPatient, i, dblLLN, dblULN) Then '/ γ-GTP(U/L)
        dblTestValue = .Cells(i, mcGTP).Value
       .Cells(i, mcGTP + 1).Value = GTP_Plus_UL(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnNa, mcNa, clPatient, i, dblLLN, dblULN) Then   '/ Na(mEq/L)
        dblTestValue = .Cells(i, mcNa).Value
       .Cells(i, mcNa + 1).Value = Na_Plus_mEqL(dblTestValue, dblULN)
       .Cells(i, mcNa + 2).Value = Na_Minus_mEqL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnK, mcK, clPatient, i, dblLLN, dblULN) Then     '/ K(mEq/L)
        dblTestValue = .Cells(i, mcK).Value
       .Cells(i, mcK + 1).Value = K_Plus_mEqL(dblTestValue, dblULN)
       .Cells(i, mcK + 2).Value = K_Minus_mEqL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnCa, mcCa, clPatient, i, dblLLN, dblULN) Then     '/ Ca(mg/dL)
        dblTestValue = .Cells(i, mcK).Value
       .Cells(i, mcCa + 1).Value = Ca_Plus_mgdL(dblTestValue, dblULN)
       .Cells(i, mcCa + 2).Value = Ca_Minus_mgdL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnIP, mcIP, clPatient, i, dblLLN, dblULN) Then     '/ IP(mg/dL)
        dblTestValue = .Cells(i, mcIP).Value
       .Cells(i, mcIP + 2).Value = IP_Minus_mgdL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnMg, mcMg, clPatient, i, dblLLN, dblULN) Then     '/ Mg(mg/dL)
        dblTestValue = .Cells(i, mcMg).Value
       .Cells(i, mcMg + 1).Value = Mg_Plus_mgdL(dblTestValue, dblULN)
       .Cells(i, mcMg + 2).Value = Mg_Minus_mgdL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnGluc, mcGluc, clPatient, i, dblLLN, dblULN) Then     '/ Gluc(mg/dL)
        dblTestValue = .Cells(i, mcGluc).Value
       .Cells(i, mcGluc + 1).Value = Gluc_Plus_mgdL(dblTestValue, dblULN)
       .Cells(i, mcGluc + 2).Value = Gluc_Minus_mgdL(dblTestValue, dblLLN)
      End If
            
      If IsReady(mcLnUPro, mcUPro, clPatient, i, dblLLN, dblULN) Then   '/ Upro
        strTestValue = .Cells(i, mcUPro).Value
       .Cells(i, mcUPro + 1).Value = UPro_Plus(strTestValue)
      End If
      
    Next
  End With
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :CalcGrade
'Argument     :None
'Return Value :None
'Date created :2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Sub CalcGrade()

  mlngMaxRow = Worksheets("Labo").UsedRange.Rows.Count
  mlngMaxRowDemog = Worksheets("Demog").UsedRange.Rows.Count
  
  If Not FirstIsReady() Then Exit Sub
  If Not DemogIsReady() Then Exit Sub
  
  Call Sheet_ApplicationOff("Labo")
  
  Call ClearSheetLabo
  
  Call SettingRefOver20
  
  Call CalcGradeMain
  
  Call Sheet_ApplicationOn("Labo")
  
End Sub


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :GetPatient
'Argument     :strCaseNum     CaseNumber
'             :dtTestday      TestDay
'Return Value :clsPatient
'Date created : 2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Function GetPatient(ByVal strCaseNum As String, ByVal dtTestday As Date) As clsPatient
  Dim i               As Long
  Dim strCurrentNum   As String
  Dim clPatient       As clsPatient
  Dim dtBirthday      As Date
  Dim lngResult       As Long
  Dim lngAgeY         As Long
  Dim lngAgeM         As Long
  
  Set GetPatient = Nothing
  Set clPatient = Nothing
  
  i = mclngDemogSttRow
  
  With Worksheets("Demog")
    For i = mclngDemogSttRow To mlngMaxRowDemog
      strCurrentNum = .Range("A" & i).Value
      If strCurrentNum = strCaseNum Then   '/ found CaseNumber
        Set clPatient = New clsPatient
        clPatient.Num = strCaseNum
        Exit For
      End If
    Next
  End With
  
  If clPatient Is Nothing Then Exit Function '/ CaseNumber None

  With Worksheets("Demog")
    dtBirthday = .Range("B" & i).Value
    clPatient.Sex = .Range("C" & i).Value
    clPatient.Cre = SetIsNumeric(.Range("D" & i).Value)
    clPatient.Hgb_gdL = SetIsNumeric(.Range("E" & i).Value)
    clPatient.Hgb_mgL = SetIsNumeric(.Range("F" & i).Value)
    clPatient.Fib = SetIsNumeric(.Range("G" & i).Value)
  End With
  
  lngResult = CalcAge(lngAgeY, lngAgeM, dtBirthday, dtTestday) '/// Get Age,Get MonthOld
  
  If Not (lngResult = 0) Then Exit Function '/ Age=None Then Exit
  
  clPatient.AgeY = lngAgeY
  clPatient.AgeM = lngAgeM

  Set GetPatient = clPatient
  
  
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :JoinKeyAgeSex  Key of Age&MonthOld&Sex("over20":Age=999,Over1:Month old =0)
'Argument     :lngAgeY
'             :lngAgeM
'             :strSex
'Return Value :KeyString
'Date created :2016/02/09 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Function JoinKeyAgeSex(ByVal lngAgeY As Long, ByVal lngAgeM As Long, ByVal strSex As String) As String
  
  If 0 < lngAgeY Then
    JoinKeyAgeSex = Format(lngAgeY, "000") & Space(1) & "00" & Space(1) & strSex
  Else
    JoinKeyAgeSex = Format(lngAgeY, "000") & Space(1) & Format(lngAgeM, "00") & Space(1) & strSex
  End If
End Function



'////////////////////////////////////////////////////////////////////////////////////////
'Name         :IsReady
'Argument     :lngRefCOL      Ref SheetCol
'             :lngLaboCOL     LaboSheetCol
'             :clPatient      Patient
'             :lngCurrentRow　CurrentRow
'             :dblLLN       　lower limit
'             :dblULN       　upper limit
'Return Value :
'Date created :2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Function IsReady(ByVal lngRefCOL As Long, ByVal lngLaboCOL As Long, ByVal clPatient As clsPatient, ByVal lngCurrentRow As Long, _
                         ByRef dblLLN As Double, ByRef dblULN As Double) As Boolean

  Dim lngRow        As Long
  Dim strValue      As String
  Dim strLLN        As String
  Dim strULN        As String
  
  IsReady = False '/ Initialization
  dblLLN = 0
  dblULN = 0
  
  strValue = Worksheets("Labo").Cells(lngCurrentRow, lngLaboCOL).Value
  If strValue = "" Then Exit Function           '/ strValue="" exit
  
  lngRow = GetRefRow(clPatient, lngRefCOL)      '/ Target Row of RefSheet
  
  If lngRow < mclngRefSttRow Then Exit Function '/ None Target Row
  
  With Worksheets("Ref")
    strLLN = .Cells(lngRow, lngRefCOL).Value
    strULN = .Cells(lngRow, lngRefCOL + 1).Value
  End With
  
  Select Case lngRefCOL
    Case mcLnUPro
      IsReady = True
      
    Case Else
      If IsNumeric(strLLN) And IsNumeric(strULN) And IsNumeric(strValue) Then
        dblLLN = strLLN
        dblULN = strULN
        IsReady = True
      End If
  End Select
  

End Function


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :GetRefRow  Row of RefSheet
'Argument     :clPatient
'             :lngTargetCol
'Return Value :RefSheet TagetRow
'Date created :2016/02/09 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Function GetRefRow(ByVal clPatient As clsPatient, ByVal lngTargetCol As Long) As Long
  Dim strAge            As String
  Dim strAgeSex         As String
  Dim strAdultAge       As String
  Dim strAdultAgeSex    As String
  Dim lngRow            As Long
  
  GetRefRow = -1
  If mcolAgeKaisou Is Nothing Then Exit Function
  If clPatient Is Nothing Then Exit Function

  strAge = JoinKeyAgeSex(clPatient.AgeY, clPatient.AgeM, "")
  strAgeSex = JoinKeyAgeSex(clPatient.AgeY, clPatient.AgeM, clPatient.Sex)
  
  strAdultAge = JoinKeyAgeSex(mclngAdult, clPatient.AgeM, "")
  strAdultAgeSex = JoinKeyAgeSex(mclngAdult, clPatient.AgeM, clPatient.Sex)
  
    
  lngRow = GetLngItem(mcolAgeKaisou, strAge)          '/ Age Only search
  If mclngRefSttRow <= lngRow Then
    If Worksheets("Ref").Cells(lngRow, lngTargetCol) <> "" Then GetRefRow = lngRow: Exit Function
  End If
  
  lngRow = GetLngItem(mcolAgeKaisou, strAgeSex)       '/ Age,Sex search
  If mclngRefSttRow <= lngRow Then
    If Worksheets("Ref").Cells(lngRow, lngTargetCol) <> "" Then GetRefRow = lngRow: Exit Function
  End If
  
  lngRow = GetLngItem(mcolAgeKaisou, strAdultAgeSex)  '/ Adult,Sex search
  If mclngRefSttRow <= lngRow Then
    If Worksheets("Ref").Cells(lngRow, lngTargetCol) <> "" Then GetRefRow = lngRow: Exit Function
  End If

  lngRow = GetLngItem(mcolAgeKaisou, strAdultAge)     '/ Adult Only search
  If mclngRefSttRow <= lngRow Then
    If Worksheets("Ref").Cells(lngRow, lngTargetCol) <> "" Then GetRefRow = lngRow: Exit Function
  End If
  
  
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :GetKaisou
'Argument     :
'Return Value :Collection 　Item:Row  Key:Age ,Month old,sex
'Date created :2016/02/08 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Function GetKaisou() As Collection
  Dim colResult   As Collection
  Dim strAgeY     As String
  Dim lngAgeY     As String
  Dim lngAgeM     As String
  Dim strSex      As String
  Dim i           As Long
  Dim strJoinKey  As String
  
  Set GetKaisou = Nothing
  
  i = mclngRefSttRow
  With Worksheets("Ref")
    Do
      strAgeY = .Range("A" & i).Value: If strAgeY = "" Then Exit Do '/ until Age=""
      lngAgeM = .Range("B" & i).Value
      strSex = .Range("C" & i).Value
      
      If IsNumeric(strAgeY) Then
        lngAgeY = strAgeY
      ElseIf strAgeY = mcstrOver20 Then
        lngAgeY = mclngAdult
      Else
        Exit Do
      End If
      
      strJoinKey = JoinKeyAgeSex(lngAgeY, lngAgeM, strSex)
      
      Call AddLngItem(colResult, i, strJoinKey)
      
      i = i + 1
    Loop
  End With
    
  Set GetKaisou = colResult
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :ClearSheetLabo
'Argument     :None
'Return Value :None
'Date created :2016/02/15 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub ClearSheetLabo()

  Call ClearSheetLaboSub(mcWBC1)
  Call ClearSheetLaboSub(mcWBC2)
  Call ClearSheetLaboSub(mcHgb1)
  Call ClearSheetLaboSub(mcHgb2)
  Call ClearSheetLaboSub(mcPLT1)
  Call ClearSheetLaboSub(mcPLT2)
  Call ClearSheetLaboSub(mcNe)
  Call ClearSheetLaboSub(mcLy)
  Call ClearSheetLaboSub(mcPT)
  Call ClearSheetLaboSub(mcAPTT)
  Call ClearSheetLaboSub(mcFib)
  Call ClearSheetLaboSub(mcALB1)
  Call ClearSheetLaboSub(mcALB2)
  Call ClearSheetLaboSub(mcCre)
  Call ClearSheetLaboSub(mcUA)
  Call ClearSheetLaboSub(mcCHO)
  Call ClearSheetLaboSub(mcTbil)
  Call ClearSheetLaboSub(mcALP)
  Call ClearSheetLaboSub(mcCPK)
  Call ClearSheetLaboSub(mcAST)
  Call ClearSheetLaboSub(mcALT)
  Call ClearSheetLaboSub(mcGTP)
  Call ClearSheetLaboSub(mcNa)
  Call ClearSheetLaboSub(mcK)
  Call ClearSheetLaboSub(mcCa)
  Call ClearSheetLaboSub(mcIP)
  Call ClearSheetLaboSub(mcMg)
  Call ClearSheetLaboSub(mcGluc)
  Call ClearSheetLaboSub(mcUPro)
  
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :ClearSheetLaboSub
'Argument     :lngCOL        TargetCol
'Return Value :None
'Date created :2016/02/15 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub ClearSheetLaboSub(ByVal lngCOL As Long)
  Worksheets("Labo").Range(Cells(mclngLaboSttRow, lngCOL + 1), Cells(mlngMaxRow, lngCOL + 2)).Value = ""
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :FirstIsReady
'Argument     :None
'Return Value :Boolean    ReadyOK then True
'Date created :2016/02/23 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Function FirstIsReady() As Boolean
  Dim strMessage    As String
  
  FirstIsReady = False
  
  strMessage = ""
  
  With Worksheets("Demog")
    If .Range("B" & mclngDemogSttRow).Value = "" Then
      strMessage = strMessage & "Demog Birthday" & vbCrLf
    End If
    If .Range("C" & mclngDemogSttRow).Value = "" Then
      strMessage = strMessage & "Demog Sex" & vbCrLf
    End If
  End With
  
  With Worksheets("Labo")
    If .Cells(mclngLaboSttRow, mcTestDay).Value = "" Then
      strMessage = strMessage & "Labo  Exam.Date(yyyy/mm/dd)" & vbCrLf
    End If
  End With
  
  If strMessage = "" Then FirstIsReady = True: Exit Function
  
  Call MsgBox("Please Input" & vbCrLf & strMessage, vbInformation Or vbOKOnly, "Input Guide")
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :DemogIsReady
'Argument     :None
'Return Value :Boolean     DemogSheetReadyOK then True(Empty 1 Only)
'Date created :2016/02/23 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Function DemogIsReady() As Boolean
  Dim i             As Long
  Dim blnIsEmpty    As Boolean
  Dim blnIsResult   As Boolean
  
  blnIsResult = True
  
  blnIsEmpty = False
  
  With Worksheets("Demog")
    For i = mclngDemogSttRow To mlngMaxRowDemog
      If (.Range("A" & i).Value = "") And (.Range("B" & i).Value <> "") Then
        If blnIsEmpty Then
          blnIsResult = False: Exit For
        Else
          blnIsEmpty = True
        End If
      End If
    Next
  End With
  
  DemogIsReady = blnIsResult
  
  If Not blnIsResult Then Call MsgBox("Please Input" & vbCrLf & "Demog Baseline ID")
  
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :SettingRefOver20
'Argument     :None
'Return Value :None
'Date created :2016/04/21 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub SettingRefOver20()
  Const clngRowOv           As Long = 35
  Const clngRowOvM          As Long = 68
  Const clngRowOvF          As Long = 101
  Dim strFacility           As String
  Dim strDate               As String
  Dim lngRows               As Long
  Dim i                     As Long
  Dim j                     As Long
  Dim lngSelectedRow        As Long
  
  With Worksheets("site")
    strFacility = .Cells(1, 2).Value
    strDate = .Cells(2, 2).Value
  End With
  
  lngSelectedRow = 0
  With Worksheets("Another_facility_date")
    lngRows = .UsedRange.Rows.Count
    Do Until lngRows < i
      i = i + 3
      
      If strFacility = .Cells(i, 1).Value And strDate = CStr(.Cells(i, 2)) Then
        lngSelectedRow = i
        Exit Do
      End If
      
    Loop
  End With
  
  If lngSelectedRow < 1 Then Exit Sub
  
  
  i = 4   '/WBC_COL of Sheet(Ref)
  j = 6   '/WBC_COL of Sheet(Another_facility_date)
  With Worksheets("Another_facility_date")
    For i = 4 To 61
      Worksheets("Ref").Cells(clngRowOv, i) = .Cells(lngSelectedRow, j)
      Worksheets("Ref").Cells(clngRowOvM, i) = .Cells(lngSelectedRow + 1, j)
      Worksheets("Ref").Cells(clngRowOvF, i) = .Cells(lngSelectedRow + 2, j)
      j = j + 1
    Next
  End With
  
End Sub
