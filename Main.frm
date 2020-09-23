VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multithreading Demo"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Vote Now !"
      Height          =   450
      Left            =   -15
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Data From Last Results"
      Height          =   450
      Left            =   3030
      TabIndex        =   5
      Top             =   1455
      Width           =   1995
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   465
      Left            =   3030
      TabIndex        =   3
      Top             =   1875
      Width           =   1995
   End
   Begin VB.CommandButton Create 
      Caption         =   "Create New Form on New Thread"
      Height          =   450
      Left            =   -90
      TabIndex        =   2
      Top             =   1890
      Width           =   1890
   End
   Begin VB.Label Process 
      Caption         =   "ProcessID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   1155
      Width           =   5010
   End
   Begin VB.Label THREAD 
      Caption         =   "ThreadID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -30
      TabIndex        =   1
      Top             =   945
      Width           =   5010
   End
   Begin VB.Label Label1 
      Caption         =   $"Main.frx":0000
      Height          =   825
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   4980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IMPORTANT ! SET THE THREADING OPTIONS TO "Thread Per Object" WHEN YOU CREATE YOUR
'OWN MULTITHREADED EXE
'RUN THIS AS A COMPILED EXE
Const SW_SHOWNORMAL = 1
Const CodeID = 24672
Dim Primes As String

Private Sub Command1_Click()
    MsgBox Primes
End Sub

Private Sub Command2_Click()
VoteNow ("http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=24672&lngWId=1")
End Sub

Private Sub Create_Click()
'   Create a new instance of the MTDemo object on a new thread
Dim Obj As MTDemoApp.MTDemo
Set Obj = CreateObject("MTDemoApp.MTDemo")
Call Obj.NewFormThread(ObjPtr(Me), Me.hwnd)
Set Obj = Nothing
End Sub

Private Sub Form_Load()
MsgBox "Make sure you are running this demo as a compiled EXE !", vbExclamation Or vbOKOnly, "Confirm !"
THREAD.Caption = "Current Thread:" + Str$(App.ThreadID)
Process.Caption = "Current Process ID:" + CStr(GetCurrentProcessId)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Check to see whether there are any other instances of this app running
    If IsOtherInstanceOpen() = False Then
        'Remove any registry entries related to this app - May not work
        'Or cause errors if any threads are running !"
        'if it does not work then perform the same operation from the command prompt"
        Shell App.Path + "\" + App.EXEName + ".exe /unregserver"
    End If
    
    'Forcefully terminate any running threads using the END statement
    'Won't work if the thread is busy
    End
    
End Sub

Private Sub Quit_Click()
Unload Me
' Explicitly use end to terminate any threads
End
End Sub

Function IsOtherInstanceOpen() As Boolean
Dim t As String, cName As String
Dim Wnd As Long
    'This function determines whether any other instances of this app are open
    'First we change the Caption of this Window to prevent
    'the FindWindow API from picking on this Window
    t = Me.Caption
    Me.Caption = ""
    
    'Now we try to find any Windows with the caption
    '"Multithreading Demo"
    Wnd = FindWindow(vbNullString, "Multithreading Demo")
    
    If Wnd <> 0 Then
    'Right ! Some other Window with the same caption is open
    'But is it a VB window ?- if it is we will assume for the
    'sake of stability that the Window belongs to another instance of this
    'App. WE will then not perform the "unregistration" of this EXE
    'in the registry and will let the other instance do it when it terminates
    
    'How do we verify whether the Window is a VB window ?
    'Simple! We call the GetClassName API to determine the Class
    'Name of the Window
    'VB Windows have class names starting with the string "Thunder"
    
        cName = String(255, " ")
        'Blank padd 255 characters
        GetClassName Wnd, cName, 255
        cName = Trim$(cName)
        'Check to see whether the text "Thunder" appears in it or not
        If InStr(1, cName, "Thunder") = 0 Then
        ' OK the Window is NOT a VB window
            IsOtherInstanceOpen = False
        Else
        'The the Window is a VB window
        'So we will assume it to be another instance of the app
            IsOtherInstanceOpen = True
        End If
    End If
    Me.Caption = t
    'Restore the caption
End Function

Sub TransferPrimes(PData As String)
'This sub will be called from the other threads for transfering data
    Primes = PData
End Sub
Sub VoteNow(URL As String)
    Dim Res As Long
    Dim TFile As String, Browser As String, Dum As String
    
    TFile = App.Path + "\test.htm"
    Open TFile For Output As #1
    Close
    Browser = String(255, " ")
    Res = FindExecutable(TFile, Dum, Browser)
    Browser = Trim$(Browser)
    
    If Len(Browser) = 0 Then
        MsgBox "Cannot find browser"
        Exit Sub
    End If
    
    Res = ShellExecute(Me.hwnd, "open", Browser, URL, sDummy, SW_SHOWNORMAL)
    If Res <= 32 Then
        MsgBox "Cannot open web page"
        Exit Sub
    End If
End Sub

