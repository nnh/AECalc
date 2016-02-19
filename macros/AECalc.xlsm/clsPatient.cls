VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'////////////////////////////////////////////////////////////////////////////////////////
'名　前：clsPatient
'説　明：Data of Patient
'作成日：2016/02/10 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////

Option Explicit

Private mstrNum       As String '/ 症例番号
Private mstrSex       As String '/ 性別
Private mlngAgeY      As Long   '/ 年齢
Private mlngAgeM      As Long   '/ 月齢
Private mdblCre       As Double '/ Cre( mg/dL)
Private mdblHgb_gdL   As Double '/ Hgb(g/dL)
Private mdblHgb_mgL   As Double '/ Hgb(mg/L)
Private mdblFib       As Double '/ フィブリノゲン(mg/dL)


Private Sub Class_Initialize()
  mstrNum = ""
  mstrSex = ""
  mlngAgeY = 0
  mlngAgeM = 0
  mdblCre = 0
  mdblHgb_gdL = 0
  mdblHgb_mgL = 0
  mdblFib = 0
End Sub

Public Property Get Sex() As String
  Sex = mstrSex
End Property

Public Property Let Sex(ByVal strSex As String)
  Select Case strSex
    Case "M", "F", "m", "f"
      mstrSex = strSex
    Case Else
      mstrSex = ""
  End Select
End Property

Public Property Get Num() As String
  Num = mstrNum
End Property

Public Property Let Num(ByVal strNum As String)
  mstrNum = strNum
End Property

Public Property Get AgeY() As Long
  AgeY = mlngAgeY
End Property

Public Property Let AgeY(ByVal lngAgeY As Long)
  mlngAgeY = lngAgeY
End Property

Public Property Get AgeM() As Long
  AgeM = mlngAgeM
End Property

Public Property Let AgeM(ByVal lngAgeM As Long)
  mlngAgeM = lngAgeM
End Property

Public Property Get Cre() As Double
  Cre = mdblCre
End Property

Public Property Let Cre(ByVal dblCre As Double)
  mdblCre = dblCre
End Property

Public Property Get Hgb_gdL() As Double
  Hgb_gdL = mdblHgb_gdL
End Property

Public Property Let Hgb_gdL(ByVal dblHgb_gdL As Double)
  mdblHgb_gdL = dblHgb_gdL
End Property

Public Property Get Hgb_mgL() As Double
  Hgb_mgL = mdblHgb_mgL
End Property

Public Property Let Hgb_mgL(ByVal dblHgb_mgL As Double)
  mdblHgb_mgL = dblHgb_mgL
End Property

Public Property Get Fib() As Double
  Fib = mdblFib
End Property

Public Property Let Fib(ByVal dblFib As Double)
  mdblFib = dblFib
End Property


