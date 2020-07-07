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
    
        Call MsgBox("OTM�t�@�C����������܂���ł���", vbCritical, "Error!")
        
        Exit Sub
    
    End If

    Dim Title As String
    
    Title = InputBox("�^�C�g������͂��ĉ�����")

    If Title = "" Then Exit Sub
    
    Dim ToFolderPath As String
    
    ToFolderPath = FolderPath & "\" & Format(Now, "YYYYMMDDHHNNSS") & "_OTMFile_" & Title
    
    Call MkDir(ToFolderPath)
    
    Dim ToFilePath As String
    
    ToFilePath = ToFolderPath & "\VbaProject.OTM"
    
    Dim FSO As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Call FSO.CopyFile(FilePath, ToFilePath)
    
    If MsgBox("OTM�t�@�C���̃R�s�[���������܂����B�t�H���_���J���܂����H", vbYesNo, "����") = vbYes Then
        
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
