VERSION 5.00
Begin VB.Form MTFORM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Form on New Thread..."
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Clear 
      Caption         =   "Clear TextBox"
      Height          =   420
      Left            =   2670
      TabIndex        =   7
      Top             =   3150
      Width           =   1995
   End
   Begin VB.TextBox pText 
      Height          =   3465
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   2400
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Primes..."
      Height          =   390
      Left            =   2670
      TabIndex        =   4
      Top             =   2775
      Width           =   2010
   End
   Begin VB.CommandButton CLOSE 
      Caption         =   "&Close"
      Height          =   345
      Left            =   2670
      TabIndex        =   3
      Top             =   2445
      Width           =   2025
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
      Height          =   645
      Left            =   2475
      TabIndex        =   8
      Top             =   3705
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   $"MTFORM.frx":0000
      Height          =   930
      Left            =   30
      TabIndex        =   5
      Top             =   1455
      Width           =   4680
   End
   Begin VB.Label Label2 
      Caption         =   "Notice the thread id is different !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   15
      TabIndex        =   2
      Top             =   915
      Width           =   4680
   End
   Begin VB.Label Thread 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15
      TabIndex        =   1
      Top             =   720
      Width           =   4680
   End
   Begin VB.Label Label1 
      Caption         =   $"MTFORM.frx":00F8
      Height          =   645
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   4680
   End
End
Attribute VB_Name = "MTFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Clear_Click()
pText.Text = ""
End Sub

Private Sub CLOSE_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
Dim min As Long, max As Long, i As Long, j As Long, t As String
min = Val(InputBox$("Enter lower range:", "Input"))
max = Val(InputBox$("Enter highler range:", "Input"))
t = Me.Caption
Me.Caption = "Performing calculations - All other forms repond !"
If max = 0 Then GoTo 20
For i = min To max
    If i = 0 Then GoTo 10
    If i = 1 Then GoTo 10
    For j = 2 To i - 1
        If i Mod j = 0 Then
            GoTo 10
        End If
    Next j
    pText.Text = pText.Text + "Prime Number:" + CStr(i) + Chr$(13) + Chr$(10)
    
10 Next i
20 Me.Caption = t
'Transfer the data
Call TransferData
End Sub

Private Sub Form_Load()
    THREAD.Caption = "New Thread Id:" + Str$(App.ThreadID)
    Process.Caption = "Current Process ID:" + CStr(GetCurrentProcessId)
End Sub


Sub TransferData()
'This sub will prepare to transfer data to the Main Thread
'Here we will transfer the contents of the textbox to the mainform
'Now first we need to retrieve a form variable from the pointer to the Main Form object
'We will use CopyMemory to do this

'Retrieve the Object Ptr stored earlier
Dim Optr As Long
Optr = GetProp(Me.hwnd, "OBJPTR")
Dim F As Form, tF As Form
CopyMemory tF, Optr, 4 '(An object variable takes 4 bytes)
Set F = tF
'Destroy TF
CopyMemory tF, 0&, 4

F.TransferPrimes (pText.Text)
'Thats'a all. the transfer is over
End Sub
