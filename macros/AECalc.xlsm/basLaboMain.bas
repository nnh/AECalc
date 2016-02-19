Attribute VB_Name = "basLaboMain"

'////////////////////////////////////////////////////////////////////////////////////////
'名　前：basLaboMain
'説　明：
'作成日：2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////


Option Explicit

'// Const
Private Const mcstrOver20         As String = "over20"
Private Const mclngAdult          As Long = 999
Private Const mclngLaboSttRow     As Long = 3     '/ LaboSheet  StartRow
Private Const mclngDemogSttRow    As Long = 2     '/ DemogSheet StartRow
Private Const mclngRefSttRow      As Long = 3     '/ Ref Sheet  StartRow

'// Labo Row
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

'// AgeSexCollection
Private mcolAgeKaisou       As collection

'////////////////////////////////////////////////////////////////////////////////////////
'名　前：KeisanGrade グレードを計算して表示する
'引　数：なし
'戻り値：なし
'作成日：2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Public Sub KeisanGrade()
  Dim strShoureiNum     As String
  Dim dtKensaday        As Date
  Dim i                 As Long
  Dim clPatient         As clsPatient
  Dim dblKensaValue     As Double
  Dim dblLLN            As Double
  Dim dblULN            As Double
  Dim dblKensaValueWBC  As Double
  Dim dblLLNWBC         As Double
  Dim dblULNWBC         As Double
  Dim strKensaValue     As String
  
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual
  Worksheets("Labo").Unprotect
  
  Call ClearSheetLabo
  Set mcolAgeKaisou = GetKaisou()
  If mcolAgeKaisou Is Nothing Then Exit Sub
  
  i = mclngLaboSttRow
  Do
    strShoureiNum = Worksheets("Labo").Cells(i, 1).Value
    
    If strShoureiNum = "" Then Exit Do                    '/ 症例番号が""まで繰り返す
    
    dtKensaday = Worksheets("Labo").Cells(i, 2).Value
    
    Set clPatient = GetPatient(strShoureiNum, dtKensaday)
    
    With Worksheets("Labo")
      If IsReady(mcLnWBC1, mcWBC1, clPatient, i, dblLLN, dblULN) Then '/ WBC(/mm3)
        dblKensaValue = .Cells(i, mcWBC1).Value
       .Cells(i, mcWBC1 + 1).Value = WBC_Plus_mm3(dblKensaValue)
       .Cells(i, mcWBC1 + 2).Value = WBC_Minus_mm3(dblKensaValue, dblLLN)
      
       dblKensaValueWBC = dblKensaValue
       dblLLNWBC = dblLLN
       If IsReady(mcLnNe, mcNe, clPatient, i, dblLLN, dblULN) Then '/ Ne(%)
        dblKensaValue = .Cells(i, mcNe).Value
        .Cells(i, mcNe + 2).Value = Ne_Minus_Per1(dblKensaValue, dblLLN, dblKensaValueWBC, dblLLNWBC)
       End If
       
       If IsReady(mcLnLy, mcLy, clPatient, i, dblLLN, dblULN) Then '/ Ly%)
        dblKensaValue = .Cells(i, mcLy).Value
        .Cells(i, mcLy + 1).Value = Ly_Plus_Per1(dblKensaValue, dblKensaValueWBC)
        .Cells(i, mcLy + 2).Value = Ly_Minus_Per1(dblKensaValue, dblLLN, dblKensaValueWBC, dblLLNWBC)
       End If
       
      End If
    
      If IsReady(mcLnWBC2, mcWBC2, clPatient, i, dblLLN, dblULN) Then  '/ WBC(WBC(10e9/L))
        dblKensaValue = .Cells(i, mcWBC2).Value
       .Cells(i, mcWBC2 + 2).Value = WBC_Minus_10e9L(dblKensaValue, dblLLN)
       
       dblKensaValueWBC = dblKensaValue
       dblLLNWBC = dblLLN
       If IsReady(mcLnNe, mcNe, clPatient, i, dblLLN, dblULN) Then      '/ Ne(%)
        dblKensaValue = .Cells(i, mcNe).Value
        .Cells(i, mcNe + 2).Value = Ne_Minus_Per2(dblKensaValue, dblLLN, dblKensaValueWBC, dblLLNWBC)
       End If
       
       If IsReady(mcLnLy, mcLy, clPatient, i, dblLLN, dblULN) Then      '/ Ly%)
        dblKensaValue = .Cells(i, mcLy).Value
        .Cells(i, mcLy + 1).Value = Ly_Plus_Per2(dblKensaValue, dblKensaValueWBC)
        .Cells(i, mcLy + 2).Value = Ly_Minus_Per2(dblKensaValue, dblLLN, dblKensaValueWBC, dblLLNWBC)
       End If
      
      End If
    
      If IsReady(mcLnHgb1, mcHgb1, clPatient, i, dblLLN, dblULN) Then '/ Hgb(g/dL)
        dblKensaValue = .Cells(i, mcHgb1).Value
       .Cells(i, mcHgb1 + 1).Value = Hgb_Plus_gdL(dblKensaValue, dblULN, clPatient.Hgb_gdL)
       .Cells(i, mcHgb1 + 2).Value = Hgb_Minus_gdL(dblKensaValue, dblLLN)
      End If
        
      If IsReady(mcLnHgb2, mcHgb2, clPatient, i, dblLLN, dblULN) Then '/ Hgb(mg/L)
        dblKensaValue = .Cells(i, mcHgb2).Value
       .Cells(i, mcHgb2 + 1).Value = Hgb_Plus_mgL(dblKensaValue, dblULN, clPatient.Hgb_mgL)
       .Cells(i, mcHgb2 + 2).Value = Hgb_Minus_mgL(dblKensaValue, dblLLN)
      End If

      If IsReady(mcLnPLT1, mcPLT1, clPatient, i, dblLLN, dblULN) Then '/ PLT(/mm3)
        dblKensaValue = .Cells(i, mcPLT1).Value
       .Cells(i, mcPLT1 + 2).Value = PLT_Minus_mm3(dblKensaValue, dblLLN)
      End If
    
      If IsReady(mcLnPLT2, mcPLT2, clPatient, i, dblLLN, dblULN) Then '/ PLT(10e9/L)
        dblKensaValue = .Cells(i, mcPLT2).Value
       .Cells(i, mcPLT2 + 2).Value = PLT_Minus_10e9L(dblKensaValue, dblLLN)
      End If
      
      If IsReady(mcLnPT, mcPT, clPatient, i, dblLLN, dblULN) Then '/ PT(PT-INR)
        dblKensaValue = .Cells(i, mcPT).Value
       .Cells(i, mcPT + 1).Value = PT_Plus_INR(dblKensaValue, dblULN)
      End If
      
      If IsReady(mcLnAPTT, mcAPTT, clPatient, i, dblLLN, dblULN) Then '/ APTT(sec)
        dblKensaValue = .Cells(i, mcAPTT).Value
       .Cells(i, mcAPTT + 1).Value = APTT_Plus_SEC(dblKensaValue, dblULN)
      End If
      
      If IsReady(mcLnFib, mcFib, clPatient, i, dblLLN, dblULN) Then '/ fib
        dblKensaValue = .Cells(i, mcFib).Value
       .Cells(i, mcFib + 2).Value = Fib_Minus_mgdL(dblKensaValue, dblLLN, clPatient.Fib)
      End If
      
      If IsReady(mcLnALB1, mcALB1, clPatient, i, dblLLN, dblULN) Then '/ ALB(g/dL)
        dblKensaValue = .Cells(i, mcALB1).Value
       .Cells(i, mcALB1 + 2).Value = ALB_Minus_gdL(dblKensaValue, dblLLN)
      End If
      
      If IsReady(mcLnALB2, mcALB2, clPatient, i, dblLLN, dblULN) Then '/ ALB(g/L)
        dblKensaValue = .Cells(i, mcALB2).Value
       .Cells(i, mcALB2 + 2).Value = ALB_Minus_gL(dblKensaValue, dblLLN)
      End If
      
      If IsReady(mcLnCre, mcCre, clPatient, i, dblLLN, dblULN) Then '/ Cre(mg/dL)
        dblKensaValue = .Cells(i, mcCre).Value
       .Cells(i, mcCre + 1).Value = Cre_Plus_mgdL(dblKensaValue, dblULN, clPatient.Cre)
       .Cells(i, mcCre + 2).Value = Cre_Plus2_mgdL(dblKensaValue, dblULN, clPatient.Cre)
      End If
      
      If IsReady(mcLnUA, mcUA, clPatient, i, dblLLN, dblULN) Then     '/ UA(mg/dL)
        dblKensaValue = .Cells(i, mcUA).Value
       .Cells(i, mcUA + 1).Value = UA_Plus_mgdL(dblKensaValue, dblULN)
      End If
     
      If IsReady(mcLnCHO, mcCHO, clPatient, i, dblLLN, dblULN) Then   '/ T-CHO(mg/dL)
        dblKensaValue = .Cells(i, mcCHO).Value
       .Cells(i, mcCHO + 1).Value = CHO_Plus_mgdL(dblKensaValue, dblULN)
      End If
     
      If IsReady(mcLnTbil, mcTbil, clPatient, i, dblLLN, dblULN) Then '/ T-Tbil(mg/dL)
        dblKensaValue = .Cells(i, mcTbil).Value
       .Cells(i, mcTbil + 1).Value = Tbil_Plus_mgdL(dblKensaValue, dblULN)
      End If
     
      If IsReady(mcLnALP, mcALP, clPatient, i, dblLLN, dblULN) Then '/ ALT(U/L)
        dblKensaValue = .Cells(i, mcALP).Value
       .Cells(i, mcALP + 1).Value = ALP_Plus_UL(dblKensaValue, dblULN)
      End If
      
      If IsReady(mcLnCPK, mcCPK, clPatient, i, dblLLN, dblULN) Then '/ CPK(U/L)
        dblKensaValue = .Cells(i, mcCPK).Value
       .Cells(i, mcCPK + 1).Value = CPK_Plus_UL(dblKensaValue, dblULN)
      End If
      
      If IsReady(mcLnAST, mcAST, clPatient, i, dblLLN, dblULN) Then '/ AST(U/L)
        dblKensaValue = .Cells(i, mcAST).Value
       .Cells(i, mcAST + 1).Value = AST_Plus_UL(dblKensaValue, dblULN)
      End If
      
      If IsReady(mcLnALT, mcALT, clPatient, i, dblLLN, dblULN) Then '/ ALT(U/L)
        dblKensaValue = .Cells(i, mcALT).Value
       .Cells(i, mcALT + 1).Value = ALT_Plus_UL(dblKensaValue, dblULN)
      End If
      
      If IsReady(mcLnGTP, mcGTP, clPatient, i, dblLLN, dblULN) Then '/ γ-GTP(U/L)
        dblKensaValue = .Cells(i, mcGTP).Value
       .Cells(i, mcGTP + 1).Value = GTP_Plus_UL(dblKensaValue, dblULN)
      End If
      
      If IsReady(mcLnNa, mcNa, clPatient, i, dblLLN, dblULN) Then   '/ Na(mEq/L)
        dblKensaValue = .Cells(i, mcNa).Value
       .Cells(i, mcNa + 1).Value = Na_Plus_mEqL(dblKensaValue, dblULN)
       .Cells(i, mcNa + 2).Value = Na_Minus_mEqL(dblKensaValue, dblLLN)
      End If
      
      If IsReady(mcLnK, mcK, clPatient, i, dblLLN, dblULN) Then     '/ K(mEq/L)
        dblKensaValue = .Cells(i, mcK).Value
       .Cells(i, mcK + 1).Value = K_Plus_mEqL(dblKensaValue, dblULN)
       .Cells(i, mcK + 2).Value = K_Minus_mEqL(dblKensaValue, dblLLN)
      End If
      
      If IsReady(mcLnCa, mcCa, clPatient, i, dblLLN, dblULN) Then     '/ Ca(mg/dL)
        dblKensaValue = .Cells(i, mcK).Value
       .Cells(i, mcCa + 1).Value = Ca_Plus_mgdL(dblKensaValue, dblULN)
       .Cells(i, mcCa + 2).Value = Ca_Minus_mgdL(dblKensaValue, dblLLN)
      End If
      
      If IsReady(mcLnIP, mcIP, clPatient, i, dblLLN, dblULN) Then     '/ IP(mg/dL)
        dblKensaValue = .Cells(i, mcIP).Value
       .Cells(i, mcIP + 2).Value = IP_Minus_mgdL(dblKensaValue, dblLLN)
      End If
      
      If IsReady(mcLnMg, mcMg, clPatient, i, dblLLN, dblULN) Then     '/ Mg(mg/dL)
        dblKensaValue = .Cells(i, mcMg).Value
       .Cells(i, mcMg + 1).Value = Mg_Plus_mgdL(dblKensaValue, dblULN)
       .Cells(i, mcMg + 2).Value = Mg_Minus_mgdL(dblKensaValue, dblLLN)
      End If
      
      If IsReady(mcLnGluc, mcGluc, clPatient, i, dblLLN, dblULN) Then     '/ Gluc(mg/dL)
        dblKensaValue = .Cells(i, mcGluc).Value
       .Cells(i, mcGluc + 1).Value = Gluc_Plus_mgdL(dblKensaValue, dblULN)
       .Cells(i, mcGluc + 2).Value = Gluc_Minus_mgdL(dblKensaValue, dblLLN)
      End If
            
      If IsReady(mcLnUPro, mcUPro, clPatient, i, dblLLN, dblULN) Then   '/ Upro
        strKensaValue = .Cells(i, mcUPro).Value
       .Cells(i, mcUPro + 1).Value = UPro_Plus(strKensaValue)
      End If

    End With
    i = i + 1 '/ NextRow
  Loop
  
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Application.Calculation = xlCalculationAutomatic
  Worksheets("Labo").Protect
End Sub


'////////////////////////////////////////////////////////////////////////////////////////
'名　前：GetPatient
'引　数：strShoureiNum  症例番号
'　　　：dtKensaday     検査日
'戻り値：clsPatient型該当患者
'作成日：2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Function GetPatient(ByVal strShoureiNum As String, ByVal dtKensaday As Date) As clsPatient
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
    Do
      strCurrentNum = .Range("A" & i).Value
      If strCurrentNum = "" Then Exit Do      '/ 症例番号が""まで繰り返す
      
      If strCurrentNum = strShoureiNum Then   '/ 症例番号見つかった
        Set clPatient = New clsPatient
        clPatient.Num = strShoureiNum
        Exit Do
      End If
      i = i + 1
    Loop
  End With
  
  If clPatient Is Nothing Then Exit Function '/ 症例番号見つからなかったなら出る

  With Worksheets("Demog")
    dtBirthday = .Range("B" & i).Value
    clPatient.Sex = .Range("C" & i).Value
    clPatient.Cre = SetIsNumeric(.Range("D" & i).Value)
    clPatient.Hgb_gdL = SetIsNumeric(.Range("E" & i).Value)
    clPatient.Hgb_mgL = SetIsNumeric(.Range("F" & i).Value)
    clPatient.Fib = SetIsNumeric(.Range("G" & i).Value)
  End With
  
  
  '/// 誕生日から年齢取得
  lngResult = CalcAge(lngAgeY, lngAgeM, dtBirthday, dtKensaday)
  
  If Not (lngResult = 0) Then Exit Function '/ 年齢取得出来なかったら出る
  
  clPatient.AgeY = lngAgeY
  clPatient.AgeM = lngAgeM

  Set GetPatient = clPatient
  
  'clPatient.KeyAgeSex = JoinKeyAgeSex(lngAgeY, lngAgeM, clPatient.Sex)
  'Debug.Print clPatient.Num & "," & clPatient.Sex & "," & clPatient.AgeY & "," & clPatient.AgeM & "," & clPatient.Cre & "," & clPatient.Hgb_mgL & "," & clPatient.Hgb_gdL & "," & clPatient.Fib
  
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'名　前：JoinKeyAgeSex 年齢 月齢 性別をスペース区切りで連結、キーを作成。
'  　　　　　　　　　　"over20" は"999"でキー　1歳以上の場合は月齢 0
'引　数：lngAgeY
'　　　：lngAgeM
'　　　：strSex
'戻り値：Keyとする文字列
'作成日：2016/02/09 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Function JoinKeyAgeSex(ByVal lngAgeY As Long, ByVal lngAgeM As Long, ByVal strSex As String) As String
  
  If 0 < lngAgeY Then
    JoinKeyAgeSex = Format(lngAgeY, "000") & Space(1) & "00" & Space(1) & strSex
  Else
    JoinKeyAgeSex = Format(lngAgeY, "000") & Space(1) & Format(lngAgeM, "00") & Space(1) & strSex
  End If
End Function



'////////////////////////////////////////////////////////////////////////////////////////
'名　前：IsReady グレード計算準備が整っていればTrueを返す
'引　数：lngRefCOL      対象Refシート列
'　　　：lngLaboCOL     対象Laboシート列
'　　　：clPatient      対象患者
'　　　：lngCurrentRow　現在行
'　　　：dblLLN       　下限
'　　　：dblULN       　上限
'戻り値：
'作成日：2016/02/10 sakaguchi
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
  If strValue = "" Then Exit Function           '/ 検査値""なら何もしない
  
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
'名　前：GetRefRow Refシートの該当ROW番号取得
'引　数：clPatient
'　　　：lngTargetCOL
'戻り値：対象となるRefシートROW番号
'作成日：2016/02/09 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Function GetRefRow(ByVal clPatient As clsPatient, ByVal lngTargetCOL As Long) As Long
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
  
    
  lngRow = GetLngItem(mcolAgeKaisou, strAge)          '/ 年齢だけでサーチ
  If mclngRefSttRow <= lngRow Then
    If Worksheets("Ref").Cells(lngRow, lngTargetCOL) <> "" Then GetRefRow = lngRow: Exit Function
  End If
  
  lngRow = GetLngItem(mcolAgeKaisou, strAgeSex)       '/ 年齢と性別でサーチ
  If mclngRefSttRow <= lngRow Then
    If Worksheets("Ref").Cells(lngRow, lngTargetCOL) <> "" Then GetRefRow = lngRow: Exit Function
  End If
  
  lngRow = GetLngItem(mcolAgeKaisou, strAdultAgeSex)  '/ 成人と性別でサーチ
  If mclngRefSttRow <= lngRow Then
    If Worksheets("Ref").Cells(lngRow, lngTargetCOL) <> "" Then GetRefRow = lngRow: Exit Function
  End If

  lngRow = GetLngItem(mcolAgeKaisou, strAdultAge)     '/ 成人だけでサーチ
  If mclngRefSttRow <= lngRow Then
    If Worksheets("Ref").Cells(lngRow, lngTargetCOL) <> "" Then GetRefRow = lngRow: Exit Function
  End If
  
  
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'名　前：GetKaisou
'引　数：
'　　　：
'戻り値：Collection 年齢性別階層コレクション　アイテム:行  Key:年齢月齢性別
'作成日：2016/02/08 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Function GetKaisou() As collection
  Dim colResult   As collection
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
      strAgeY = .Range("A" & i).Value: If strAgeY = "" Then Exit Do '/ 年齢が""まで繰り返す
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
'名　前：ClearSheetLabo
'引　数：なし
'戻り値：なし
'作成日：2016/02/15 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub ClearSheetLabo()
  Dim lngMaxRow  As Long
  
  lngMaxRow = Worksheets("Labo").UsedRange.Rows.Count

  Call ClearSheetLaboSub(lngMaxRow, mcWBC1)
  Call ClearSheetLaboSub(lngMaxRow, mcWBC2)
  Call ClearSheetLaboSub(lngMaxRow, mcHgb1)
  Call ClearSheetLaboSub(lngMaxRow, mcHgb2)
  Call ClearSheetLaboSub(lngMaxRow, mcPLT1)
  Call ClearSheetLaboSub(lngMaxRow, mcPLT2)
  Call ClearSheetLaboSub(lngMaxRow, mcNe)
  Call ClearSheetLaboSub(lngMaxRow, mcLy)
  Call ClearSheetLaboSub(lngMaxRow, mcPT)
  Call ClearSheetLaboSub(lngMaxRow, mcAPTT)
  Call ClearSheetLaboSub(lngMaxRow, mcFib)
  Call ClearSheetLaboSub(lngMaxRow, mcALB1)
  Call ClearSheetLaboSub(lngMaxRow, mcALB2)
  Call ClearSheetLaboSub(lngMaxRow, mcCre)
  Call ClearSheetLaboSub(lngMaxRow, mcUA)
  Call ClearSheetLaboSub(lngMaxRow, mcCHO)
  Call ClearSheetLaboSub(lngMaxRow, mcTbil)
  Call ClearSheetLaboSub(lngMaxRow, mcALP)
  Call ClearSheetLaboSub(lngMaxRow, mcCPK)
  Call ClearSheetLaboSub(lngMaxRow, mcAST)
  Call ClearSheetLaboSub(lngMaxRow, mcALT)
  Call ClearSheetLaboSub(lngMaxRow, mcGTP)
  Call ClearSheetLaboSub(lngMaxRow, mcNa)
  Call ClearSheetLaboSub(lngMaxRow, mcK)
  Call ClearSheetLaboSub(lngMaxRow, mcCa)
  Call ClearSheetLaboSub(lngMaxRow, mcIP)
  Call ClearSheetLaboSub(lngMaxRow, mcMg)
  Call ClearSheetLaboSub(lngMaxRow, mcGluc)
  Call ClearSheetLaboSub(lngMaxRow, mcUPro)
  
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'名　前：ClearSheetLaboSub
'引　数：lngMaxRow　最大行
'　　　：lngCOL     対象項目列
'戻り値：なし
'作成日：2016/02/15 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub ClearSheetLaboSub(ByVal lngMaxRow As Long, ByVal lngCOL As Long)
  Worksheets("Labo").Range(Cells(mclngLaboSttRow, lngCOL + 1), Cells(lngMaxRow, lngCOL + 2)).Value = ""
End Sub
