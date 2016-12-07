Attribute VB_Name = "basLaboMain"
'////////////////////////////////////////////////////////////////////////////////////////
'Name         :basLaboMain
'Explanation  :
'Date created : 2016/02/10 sakaguchi
'             : 2016/04/21 sakaguchi (Add SettingRefOver20)
'　　　　　　 : 2016/12/07 sakaguchi (Add WBC_Plus_10e9L)
'////////////////////////////////////////////////////////////////////////////////////////

Option Explicit

'// Public Const
Public Const gclngLaboSttRow     As Long = 3     '/ LaboSheet  StartRow

'// Labo Row
Public Const gcCaseNo      As Long = 1
Public Const gcTestDay     As Long = 2
Public Const gcWBC1        As Long = 3
Public Const gcWBC2        As Long = 6
Public Const gcHgb1        As Long = 9
Public Const gcHgb2        As Long = 12
Public Const gcPLT1        As Long = 15
Public Const gcPLT2        As Long = 18
Public Const gcNe          As Long = 21
Public Const gcLy          As Long = 24
Public Const gcPT          As Long = 27
Public Const gcAPTT        As Long = 30
Public Const gcFib         As Long = 33
Public Const gcALB1        As Long = 36
Public Const gcALB2        As Long = 39
Public Const gcCre         As Long = 42
Public Const gcUA          As Long = 45
Public Const gcCHO         As Long = 48
Public Const gcTbil        As Long = 51
Public Const gcALP         As Long = 54
Public Const gcCPK         As Long = 57
Public Const gcAST         As Long = 60
Public Const gcALT         As Long = 63
Public Const gcGTP         As Long = 66
Public Const gcNa          As Long = 69
Public Const gcK           As Long = 72
Public Const gcCa          As Long = 75
Public Const gcIP          As Long = 78
Public Const gcMg          As Long = 81
Public Const gcGluc        As Long = 84
Public Const gcUPro        As Long = 87

'// Private Const
Private Const mcstrOver20         As String = "over20"
Private Const mclngAdult          As Long = 999
Private Const mclngDemogSttRow    As Long = 2     '/ DemogSheet StartRow
Private Const mclngRefSttRow      As Long = 3     '/ Ref Sheet  StartRow

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
    For i = gclngLaboSttRow To mlngMaxRow
      strCaseNum = Worksheets("Labo").Cells(i, gcCaseNo).Value
        
      dtTestday = Worksheets("Labo").Cells(i, gcTestDay).Value
    
      Set clPatient = GetPatient(strCaseNum, dtTestday)
    
      If IsReady(mcLnWBC1, gcWBC1, clPatient, i, dblLLN, dblULN) Then '/ WBC(/mm3)
        dblTestValue = .Cells(i, gcWBC1).Value
       .Cells(i, gcWBC1 + 1).Value = WBC_Plus_mm3(dblTestValue)
       .Cells(i, gcWBC1 + 2).Value = WBC_Minus_mm3(dblTestValue, dblLLN)
      
       dblTestValueWBC = dblTestValue
       dblLLNWBC = dblLLN
       If IsReady(mcLnNe, gcNe, clPatient, i, dblLLN, dblULN) Then    '/ Ne(%)
        dblTestValue = .Cells(i, gcNe).Value
        .Cells(i, gcNe + 2).Value = Ne_Minus_Per1(dblTestValue, dblLLN, dblTestValueWBC, dblLLNWBC)
       End If
       
       If IsReady(mcLnLy, gcLy, clPatient, i, dblLLN, dblULN) Then    '/ Ly(%)
        dblTestValue = .Cells(i, gcLy).Value
        .Cells(i, gcLy + 1).Value = Ly_Plus_Per1(dblTestValue, dblTestValueWBC)
        .Cells(i, gcLy + 2).Value = Ly_Minus_Per1(dblTestValue, dblLLN, dblTestValueWBC, dblLLNWBC)
       End If
       
      End If
    
      If IsReady(mcLnWBC2, gcWBC2, clPatient, i, dblLLN, dblULN) Then  '/ WBC(10e9/L)
        dblTestValue = .Cells(i, gcWBC2).Value
        .Cells(i, gcWBC2 + 1).Value = WBC_Plus_10e9L(dblTestValue)
        .Cells(i, gcWBC2 + 2).Value = WBC_Minus_10e9L(dblTestValue, dblLLN)
       
       dblTestValueWBC = dblTestValue
       dblLLNWBC = dblLLN
       If IsReady(mcLnNe, gcNe, clPatient, i, dblLLN, dblULN) Then      '/ Ne(%)
        dblTestValue = .Cells(i, gcNe).Value
        .Cells(i, gcNe + 2).Value = Ne_Minus_Per2(dblTestValue, dblLLN, dblTestValueWBC, dblLLNWBC)
       End If
       
       If IsReady(mcLnLy, gcLy, clPatient, i, dblLLN, dblULN) Then      '/ Ly%)
        dblTestValue = .Cells(i, gcLy).Value
        .Cells(i, gcLy + 1).Value = Ly_Plus_Per2(dblTestValue, dblTestValueWBC)
        .Cells(i, gcLy + 2).Value = Ly_Minus_Per2(dblTestValue, dblLLN, dblTestValueWBC, dblLLNWBC)
       End If
      
      End If
    
      If IsReady(mcLnHgb1, gcHgb1, clPatient, i, dblLLN, dblULN) Then '/ Hgb(g/dL)
        dblTestValue = .Cells(i, gcHgb1).Value
       .Cells(i, gcHgb1 + 1).Value = Hgb_Plus_gdL(dblTestValue, dblULN, clPatient.Hgb_gdL)
       .Cells(i, gcHgb1 + 2).Value = Hgb_Minus_gdL(dblTestValue, dblLLN)
      End If
        
      If IsReady(mcLnHgb2, gcHgb2, clPatient, i, dblLLN, dblULN) Then '/ Hgb(mg/L)
        dblTestValue = .Cells(i, gcHgb2).Value
       .Cells(i, gcHgb2 + 1).Value = Hgb_Plus_mgL(dblTestValue, dblULN, clPatient.Hgb_mgL)
       .Cells(i, gcHgb2 + 2).Value = Hgb_Minus_mgL(dblTestValue, dblLLN)
      End If

      If IsReady(mcLnPLT1, gcPLT1, clPatient, i, dblLLN, dblULN) Then '/ PLT(/mm3)
        dblTestValue = .Cells(i, gcPLT1).Value
       .Cells(i, gcPLT1 + 2).Value = PLT_Minus_mm3(dblTestValue, dblLLN)
      End If
    
      If IsReady(mcLnPLT2, gcPLT2, clPatient, i, dblLLN, dblULN) Then '/ PLT(10e9/L)
        dblTestValue = .Cells(i, gcPLT2).Value
       .Cells(i, gcPLT2 + 2).Value = PLT_Minus_10e9L(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnPT, gcPT, clPatient, i, dblLLN, dblULN) Then '/ PT(PT-INR)
        dblTestValue = .Cells(i, gcPT).Value
       .Cells(i, gcPT + 1).Value = PT_Plus_INR(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnAPTT, gcAPTT, clPatient, i, dblLLN, dblULN) Then '/ APTT(sec)
        dblTestValue = .Cells(i, gcAPTT).Value
       .Cells(i, gcAPTT + 1).Value = APTT_Plus_SEC(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnFib, gcFib, clPatient, i, dblLLN, dblULN) Then '/ fib
        dblTestValue = .Cells(i, gcFib).Value
       .Cells(i, gcFib + 2).Value = Fib_Minus_mgdL(dblTestValue, dblLLN, clPatient.Fib)
      End If
      
      If IsReady(mcLnALB1, gcALB1, clPatient, i, dblLLN, dblULN) Then '/ ALB(g/dL)
        dblTestValue = .Cells(i, gcALB1).Value
       .Cells(i, gcALB1 + 2).Value = ALB_Minus_gdL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnALB2, gcALB2, clPatient, i, dblLLN, dblULN) Then '/ ALB(g/L)
        dblTestValue = .Cells(i, gcALB2).Value
       .Cells(i, gcALB2 + 2).Value = ALB_Minus_gL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnCre, gcCre, clPatient, i, dblLLN, dblULN) Then '/ Cre(mg/dL)
        dblTestValue = .Cells(i, gcCre).Value
       .Cells(i, gcCre + 1).Value = Cre_Plus_mgdL(dblTestValue, dblULN, clPatient.Cre)
       .Cells(i, gcCre + 2).Value = Cre_Plus2_mgdL(dblTestValue, dblULN, clPatient.Cre)
      End If
      
      If IsReady(mcLnUA, gcUA, clPatient, i, dblLLN, dblULN) Then     '/ UA(mg/dL)
        dblTestValue = .Cells(i, gcUA).Value
       .Cells(i, gcUA + 1).Value = UA_Plus_mgdL(dblTestValue, dblULN)
      End If
     
      If IsReady(mcLnCHO, gcCHO, clPatient, i, dblLLN, dblULN) Then   '/ T-CHO(mg/dL)
        dblTestValue = .Cells(i, gcCHO).Value
       .Cells(i, gcCHO + 1).Value = CHO_Plus_mgdL(dblTestValue, dblULN)
      End If
     
      If IsReady(mcLnTbil, gcTbil, clPatient, i, dblLLN, dblULN) Then '/ T-Tbil(mg/dL)
        dblTestValue = .Cells(i, gcTbil).Value
       .Cells(i, gcTbil + 1).Value = Tbil_Plus_mgdL(dblTestValue, dblULN)
      End If
     
      If IsReady(mcLnALP, gcALP, clPatient, i, dblLLN, dblULN) Then '/ ALT(U/L)
        dblTestValue = .Cells(i, gcALP).Value
       .Cells(i, gcALP + 1).Value = ALP_Plus_UL(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnCPK, gcCPK, clPatient, i, dblLLN, dblULN) Then '/ CPK(U/L)
        dblTestValue = .Cells(i, gcCPK).Value
       .Cells(i, gcCPK + 1).Value = CPK_Plus_UL(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnAST, gcAST, clPatient, i, dblLLN, dblULN) Then '/ AST(U/L)
        dblTestValue = .Cells(i, gcAST).Value
       .Cells(i, gcAST + 1).Value = AST_Plus_UL(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnALT, gcALT, clPatient, i, dblLLN, dblULN) Then '/ ALT(U/L)
        dblTestValue = .Cells(i, gcALT).Value
       .Cells(i, gcALT + 1).Value = ALT_Plus_UL(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnGTP, gcGTP, clPatient, i, dblLLN, dblULN) Then '/ γ-GTP(U/L)
        dblTestValue = .Cells(i, gcGTP).Value
       .Cells(i, gcGTP + 1).Value = GTP_Plus_UL(dblTestValue, dblULN)
      End If
      
      If IsReady(mcLnNa, gcNa, clPatient, i, dblLLN, dblULN) Then   '/ Na(mEq/L)
        dblTestValue = .Cells(i, gcNa).Value
       .Cells(i, gcNa + 1).Value = Na_Plus_mEqL(dblTestValue, dblULN)
       .Cells(i, gcNa + 2).Value = Na_Minus_mEqL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnK, gcK, clPatient, i, dblLLN, dblULN) Then     '/ K(mEq/L)
        dblTestValue = .Cells(i, gcK).Value
       .Cells(i, gcK + 1).Value = K_Plus_mEqL(dblTestValue, dblULN)
       .Cells(i, gcK + 2).Value = K_Minus_mEqL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnCa, gcCa, clPatient, i, dblLLN, dblULN) Then     '/ Ca(mg/dL)
        dblTestValue = .Cells(i, gcCa).Value
       .Cells(i, gcCa + 1).Value = Ca_Plus_mgdL(dblTestValue, dblULN)
       .Cells(i, gcCa + 2).Value = Ca_Minus_mgdL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnIP, gcIP, clPatient, i, dblLLN, dblULN) Then     '/ IP(mg/dL)
        dblTestValue = .Cells(i, gcIP).Value
       .Cells(i, gcIP + 2).Value = IP_Minus_mgdL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnMg, gcMg, clPatient, i, dblLLN, dblULN) Then     '/ Mg(mg/dL)
        dblTestValue = .Cells(i, gcMg).Value
       .Cells(i, gcMg + 1).Value = Mg_Plus_mgdL(dblTestValue, dblULN)
       .Cells(i, gcMg + 2).Value = Mg_Minus_mgdL(dblTestValue, dblLLN)
      End If
      
      If IsReady(mcLnGluc, gcGluc, clPatient, i, dblLLN, dblULN) Then     '/ Gluc(mg/dL)
        dblTestValue = .Cells(i, gcGluc).Value
       .Cells(i, gcGluc + 1).Value = Gluc_Plus_mgdL(dblTestValue, dblULN)
       .Cells(i, gcGluc + 2).Value = Gluc_Minus_mgdL(dblTestValue, dblLLN)
      End If
            
      If IsReady(mcLnUPro, gcUPro, clPatient, i, dblLLN, dblULN) Then   '/ Upro
        strTestValue = .Cells(i, gcUPro).Value
       .Cells(i, gcUPro + 1).Value = UPro_Plus(strTestValue)
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

  Dim lngROW        As Long
  Dim strValue      As String
  Dim strLLN        As String
  Dim strULN        As String
  
  IsReady = False '/ Initialization
  dblLLN = 0
  dblULN = 0
  
  strValue = Worksheets("Labo").Cells(lngCurrentRow, lngLaboCOL).Value
  If strValue = "" Then Exit Function           '/ strValue="" exit
  
  lngROW = GetRefRow(clPatient, lngRefCOL)      '/ Target Row of RefSheet
  
  If lngROW < mclngRefSttRow Then Exit Function '/ None Target Row
  
  With Worksheets("Ref")
    strLLN = .Cells(lngROW, lngRefCOL).Value
    strULN = .Cells(lngROW, lngRefCOL + 1).Value
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
  Dim lngROW            As Long
  
  GetRefRow = -1
  If mcolAgeKaisou Is Nothing Then Exit Function
  If clPatient Is Nothing Then Exit Function

  strAge = JoinKeyAgeSex(clPatient.AgeY, clPatient.AgeM, "")
  strAgeSex = JoinKeyAgeSex(clPatient.AgeY, clPatient.AgeM, clPatient.Sex)
  
  strAdultAge = JoinKeyAgeSex(mclngAdult, clPatient.AgeM, "")
  strAdultAgeSex = JoinKeyAgeSex(mclngAdult, clPatient.AgeM, clPatient.Sex)
  
    
  lngROW = GetLngItem(mcolAgeKaisou, strAge)          '/ Age Only search
  If mclngRefSttRow <= lngROW Then
    If Worksheets("Ref").Cells(lngROW, lngTargetCol) <> "" Then GetRefRow = lngROW: Exit Function
  End If
  
  lngROW = GetLngItem(mcolAgeKaisou, strAgeSex)       '/ Age,Sex search
  If mclngRefSttRow <= lngROW Then
    If Worksheets("Ref").Cells(lngROW, lngTargetCol) <> "" Then GetRefRow = lngROW: Exit Function
  End If
  
  lngROW = GetLngItem(mcolAgeKaisou, strAdultAgeSex)  '/ Adult,Sex search
  If mclngRefSttRow <= lngROW Then
    If Worksheets("Ref").Cells(lngROW, lngTargetCol) <> "" Then GetRefRow = lngROW: Exit Function
  End If

  lngROW = GetLngItem(mcolAgeKaisou, strAdultAge)     '/ Adult Only search
  If mclngRefSttRow <= lngROW Then
    If Worksheets("Ref").Cells(lngROW, lngTargetCol) <> "" Then GetRefRow = lngROW: Exit Function
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

  Call ClearSheetLaboSub(gcWBC1)
  Call ClearSheetLaboSub(gcWBC2)
  Call ClearSheetLaboSub(gcHgb1)
  Call ClearSheetLaboSub(gcHgb2)
  Call ClearSheetLaboSub(gcPLT1)
  Call ClearSheetLaboSub(gcPLT2)
  Call ClearSheetLaboSub(gcNe)
  Call ClearSheetLaboSub(gcLy)
  Call ClearSheetLaboSub(gcPT)
  Call ClearSheetLaboSub(gcAPTT)
  Call ClearSheetLaboSub(gcFib)
  Call ClearSheetLaboSub(gcALB1)
  Call ClearSheetLaboSub(gcALB2)
  Call ClearSheetLaboSub(gcCre)
  Call ClearSheetLaboSub(gcUA)
  Call ClearSheetLaboSub(gcCHO)
  Call ClearSheetLaboSub(gcTbil)
  Call ClearSheetLaboSub(gcALP)
  Call ClearSheetLaboSub(gcCPK)
  Call ClearSheetLaboSub(gcAST)
  Call ClearSheetLaboSub(gcALT)
  Call ClearSheetLaboSub(gcGTP)
  Call ClearSheetLaboSub(gcNa)
  Call ClearSheetLaboSub(gcK)
  Call ClearSheetLaboSub(gcCa)
  Call ClearSheetLaboSub(gcIP)
  Call ClearSheetLaboSub(gcMg)
  Call ClearSheetLaboSub(gcGluc)
  Call ClearSheetLaboSub(gcUPro)
  
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :ClearSheetLaboSub
'Argument     :lngCOL        TargetCol
'Return Value :None
'Date created :2016/02/15 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub ClearSheetLaboSub(ByVal lngCOL As Long)
  Worksheets("Labo").Range(Cells(gclngLaboSttRow, lngCOL + 1), Cells(mlngMaxRow, lngCOL + 2)).Value = ""
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
    If .Cells(gclngLaboSttRow, gcTestDay).Value = "" Then
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
