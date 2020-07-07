Attribute VB_Name = "OLOTMFile"
Option Explicit

Public Sub CopyOTMFile()

    Dim UserName As String
    
    UserName = GetUserName
    
    Dim FolderPath As String
        
    FolderPath = "C:\Users\" & UserName & "\AppData\Roaming\Microsoft\Outlook"
    
    Const FileName As String = "VbaProject.OTM"

    Dim FilePath As String
    
    FilePath = FolderPath & "\" & FileName

    If Dir(FilePath) = "" Then
    
        Call MsgBox("OTMファイルが見つかりませんでした", vbCritical, "Error!")
        
        Exit Sub
    
    End If

    Dim Title As String
    
    Title = InputBox("タイトルを入力して下さい")

    If Title = "" Then Exit Sub
    
    Dim ToFolderPath As String
    
    ToFolderPath = FolderPath & "\" & Format(Now, "YYYYMMDDHHNNSS") & "_OTMFile_" & Title
    
    Call MkDir(ToFolderPath)
    
    Dim ToFilePath As String
    
    ToFilePath = ToFolderPath & "\VbaProject.OTM"
    
    Dim FSO As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Call FSO.CopyFile(FilePath, ToFilePath)
    
    If MsgBox("OTMファイルのコピーが完了しました。フォルダを開きますか？", vbYesNo, "完了") = vbYes Then
        
        Call Shell("C:\Windows\explorer.exe " & ToFolderPath, vbNormalFocus)
    
    End If
    
End Sub
Public Sub OpenOTMFilePath()

    Dim UserName As String
    
    UserName = GetUserName
    
    Dim FolderPath As String
        
    FolderPath = "C:\Users\" & UserName & "\AppData\Roaming\Microsoft\Outlook"
    
    Call Shell("C:\Windows\Explorer.exe " & FolderPath, vbNormalFocus)
    
End Sub
Private Function GetUserName() As String

    GetUserName = CreateObject("WScript.Network").UserName
    
End Function
