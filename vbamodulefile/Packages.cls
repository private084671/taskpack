VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Packages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private ParentPath As String

Private Sub GetParentPath()

    Dim UserName As String
    
    UserName = WshNetwork.GetUserName
    
    If UserName = "" Then Exit Sub 'UserNameが取得できず怪しいパスが生成されるのを防止
    
    ParentPath = "C:\Users\" & UserName & "\TaskPack\Packages"

End Sub

Private Function ExistsParentPath() As Boolean

    If Dir(ParentPath, vbDirectory) = "." Then Exit Function

    If Dir(ParentPath, vbDirectory) <> "" Then ExistsParentPath = True

End Function

Private Sub mkParentPath()

    On Error GoTo FailedMkPath

    Call MkDir(ParentPath)

FailedMkPath:

End Sub

Private Sub Class_Initialize()

    Call GetParentPath

    If Not ExistsParentPath Then Call mkParentPath
    
End Sub
