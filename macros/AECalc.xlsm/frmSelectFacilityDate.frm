VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectFacilityDate 
   Caption         =   "Select Facility Date"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   OleObjectBlob   =   "frmSelectFacilityDate.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelectFacilityDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////////////////////////////
'Name         : frmSelectFacilityDate
'Explanation  : Select　Data of Facility
'Date created : 2016/04/19 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Option Explicit

'// variable
Private mcolFacility      As Collection   '/ Choices
Private mstrFacility      As String       '/ Selected Facility
Private mstrDate          As String       '/ Selected Date

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdDecision_Click()
  Call DecisionMain
End Sub

Private Sub lstDate_Click()
  Call SettingEnabledDecision
End Sub

Private Sub UserForm_Initialize()
   
  Set mcolFacility = Nothing
  mstrFacility = ""
  mstrDate = ""
 
  Call InitializeList
  
End Sub


Public Property Get SelectedFacility() As String
  SelectedFacility = mstrFacility
End Property


Public Property Let SelectedFacility(ByVal strFacilityValue As String)
  mstrFacility = strFacilityValue
End Property

Public Property Get SelectedDate() As String
  SelectedDate = mstrDate
End Property

Public Property Let SelectedDate(ByVal strDateValue As String)
  mstrDate = strDateValue
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
  End If
End Sub


'////////////////////////////////////////////////////////////////////////////////////////
'Name         :InitializeList
'Argument     :None
'Return Value :None
'Date created :2016/04/19 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub InitializeList()
  Call SettingFacilityList
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :SettingFacilityList
'Argument     :None
'Return Value :None
'Date created :2016/04/19 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub SettingFacilityList()
  Dim lngRows       As Long
  Dim i             As Long
  Dim strBufName    As String
  Dim strBufDate    As String
  Dim clFacility    As clsFacility

  lstFacility.Clear
  lstDate.Clear
  
  lngRows = Worksheets("Another_facility_date").UsedRange.Rows.Count
  strBufName = ""
  strBufDate = ""
  i = 0

  With Worksheets("Another_facility_date")
    Do Until lngRows < i
        
      i = i + 3
      strBufDate = CStr(.Cells(i, 2))
      
      If strBufName <> .Cells(i, 1).Value Then
        If 3 < i Then mcolFacility.Add clFacility
      
        strBufName = .Cells(i, 1).Value
        
        If strBufName = "" Then Exit Do
        lstFacility.AddItem strBufName
       
        If mcolFacility Is Nothing Then Set mcolFacility = New Collection
        Set clFacility = New clsFacility
        clFacility.Name = strBufName
        clFacility.colDate = New Collection
        
      End If
      
      On Error Resume Next
      clFacility.colDate.Add strBufDate, strBufDate
      On Error GoTo 0

    Loop
  End With
  
End Sub



'////////////////////////////////////////////////////////////////////////////////////////
'Name         :SettingDateList
'Argument     :None
'Return Value :None
'Date created :2016/04/19 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub SettingDateList()
  Dim clFacility  As clsFacility
  Dim lngItem     As Long
  Dim varBuf      As Variant
  
  lstDate.Clear
  lngItem = lstFacility.ListIndex + 1
  
  On Error Resume Next
  Set clFacility = mcolFacility.Item(lngItem)
  On Error GoTo 0
  If Not (clFacility Is Nothing) Then
    For Each varBuf In clFacility.colDate
      lstDate.AddItem varBuf
    Next
  End If
  
  Call SettingEnabledDecision
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :SettingEnabledDecision
'Argument     :None
'Return Value :None
'Date created :2016/04/21 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub SettingEnabledDecision()
  cmdDecision.Enabled = False
  If lstFacility.ListCount < 1 Then Exit Sub
  If lstFacility.ListIndex < 0 Then Exit Sub
  If lstDate.ListCount < 1 Then Exit Sub
  If lstDate.ListIndex < 0 Then Exit Sub
  cmdDecision.Enabled = True
End Sub

Private Sub UserForm_Activate()
  Dim i As Long
  Dim j As Long
  
  For i = 0 To lstFacility.ListCount - 1
    If lstFacility.List(i) = mstrFacility Then
      lstFacility.ListIndex = i
      Call SettingDateList
      
      For j = 0 To lstDate.ListCount - 1
        If lstDate.List(j) = mstrDate Then
          lstDate.ListIndex = j
          Exit For
        End If
      Next
      
      Exit For
    End If
  Next

  Call SettingEnabledDecision
End Sub

Private Sub lstDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call DecisionMain
End Sub

Private Sub lstFacility_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call SettingDateList
End Sub

Private Sub lstFacility_Click()
  Call SettingDateList
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'Name         :DecisionMain
'Argument     :None
'Return Value :None
'Date created :2016/04/21 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub DecisionMain()
  mstrFacility = lstFacility.List(lstFacility.ListIndex)
  mstrDate = lstDate.List(lstDate.ListIndex)
  Me.Hide
End Sub
