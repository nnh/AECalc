VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFacility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////////////////////////////
'Name         : clsFacility
'Explanation  : Data of Facility
'Date created : 2016/04/19 sakaguchi
'////////////////////////////////////////////////////////////////////////////////////////

Option Explicit

Private mstrName        As String     '/ FacilityName
Private mcolDate        As Collection '/ DateCollection

Public Property Get Name() As String
  Name = mstrName
End Property

Public Property Let Name(ByVal strNameValue As String)
  mstrName = strNameValue
End Property


Public Property Get colDate() As Collection
  Set colDate = mcolDate
End Property

Public Property Let colDate(ByVal colDateValue As Collection)
  Set mcolDate = colDateValue
End Property

Private Sub Class_Initialize()
  mstrName = ""
  Set mcolDate = Nothing
End Sub

