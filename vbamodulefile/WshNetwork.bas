Attribute VB_Name = "WshNetwork"
Option Explicit

Public Function GetUserName() As String

    GetUserName = GetNetWork.UserName
    
End Function

Private Function GetNetWork() As Object

    Dim NetWork As Object
    
    Set GetNetWork = CreateObject("Wscript.NetWork")

End Function
