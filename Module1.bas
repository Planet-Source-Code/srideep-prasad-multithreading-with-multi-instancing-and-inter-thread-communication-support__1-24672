Attribute VB_Name = "Module1"
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long



'MAKE SURE TO SET THE PROJECT PROPERTIES TO STANDALONE
'EXE AND NOT ActiveX SERVER
'RUN THIS DEMO ONLY AS A COMPILED EXE
'MAKE SURE THE THREADING OPTION IS "THREAD PER OBJECT"
Sub Main()
    Dim ProcessID As Long, curProcessID As Long
    hwnd = FindWindow(vbNullString, "Multithreading Demo")
'   This routine is called whenever a new object (on a new thread)
'   is created. Therefore we use the FindWindow API to check
'   whether the main form is loaded or not.
'   If not we load it
    
    If hwnd <> 0 Then
    'We perform an additional check here since the window
    'with the title "Multithreading Demo" can be any other window
    'Other than that of this app !
    'If this check is not present then we may prevent
    ' the loading of our app just because another Window
    'with the same title happens to be open
    'Also we need to able to start multiple instances of our app
    'which will not be possible without this step !
    
    'Get the ProcessID of the Windows identified with the hwnd Handle
    'returned
    'This we compare with the ProcessID of our app to see
    'whether the supposed "Main Window" is our app's Window or not
    'Also we compare the processIds to allow users to start multiple
    'instances of our app !
        GetWindowThreadProcessId hwnd, ProcessID
        curProcessID = GetCurrentProcessId
        'Compare both process ids
        
        If curProcessID <> ProcessID Then
'        Main form not loaded, so load it
            Dim Frm As New Form1
                Frm.Show
            Set Frm = Nothing

        End If
    End If
        
        
             
   If hwnd = 0 Then
'       A Window with such a title does not exist
'       So no problem ! Just directly load it
            Dim Frm2 As New Form1
                Frm2.Show
            Set Frm2 = Nothing
    End If
        'Otherwise do nothing and let the secondary objects be created
End Sub


