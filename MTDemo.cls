VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MTDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Sub NewFormThread(OPtr As Long, MainFormHandle As Long)
'The interesting part is here
'We store the WindowHandle of the MainWindow and a pointer
'to the Main Window Form by using the SetProp function against
'The Hwnd of the New Form as a virtual property
    Dim Frm As New MTFORM
    Frm.Show
    SetProp Frm.hwnd, "MAINHWND", MainFormHandle
    SetProp Frm.hwnd, "OBJPTR", OPtr
    Set Frm = Nothing
End Sub
