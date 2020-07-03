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

    Private PackagesPath As String

Private Sub CreatePackagesPath()

    Dim NetWork As Object
    
    Set NetWork = CreateObject("Wscript.NetWork")
    
    PackagesPath = "C:\Users\" & NetWork.UserName & "\TaskPack\Packages"

End Sub

Private Sub Class_Initialize()

    Call CreatePackagesPath

End Sub
