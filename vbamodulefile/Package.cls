VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "package"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private mPath As String
    
Public Property Let Path(ByRef FullPath As String)

    mPath = FullPath

End Property

Public Property Get Path() As String

    Path = mPath

End Property

Public Property Get Exists() As Boolean

    

End Property

Public Sub AddTask()

End Sub

Public Sub ChangeTitle()

End Sub

Public Sub Delete()

End Sub


