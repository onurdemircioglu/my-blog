VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormOpenFiles 
   Caption         =   "Open Files or Folders"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11955
   OleObjectBlob   =   "UserFormOpenFiles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormOpenFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    With Me
        .Height = 175
        .Width = 610
    End With
End Sub

Private Sub cmdExit_Click()
    Unload UserFormOpenFiles
End Sub

Private Sub txtLink_Change()
    With Me
        .OptBtnFile.Value = False
        .OptBtnFolder.Value = False
        .OptBtnMoreThanOneFile.Value = False
        .lstFileList.Clear
    End With
    
    If Trim(Me.txtLink) <> "" Then
        With Me
            .Frame1.Enabled = True
        End With
    End If
End Sub

Private Sub OptBtnFile_Click()
    With UserFormOpenFiles
        .lstFileList.Visible = False
        .OptBtnFileRead.Value = False
        .OptBtnFileWrite.Value = False
        .Frame2.Enabled = True
    End With
    
    If Right(Me.txtLink.Value, 5) = ".xlsx" Or Right(Me.txtLink.Value, 5) = ".xlsm" Or Right(Me.txtLink.Value, 4) = ".xls" Then

    Else
        Me.txtLink.Value = Null
        MsgBox "File should be an Excel file."
    End If
End Sub

Private Sub OptBtnFolder_Click()
    With UserFormOpenFiles
        .lstFileList.Visible = False
        .OptBtnFileRead.Value = False
        .OptBtnFileWrite.Value = False
        .Frame2.Enabled = False
    End With
End Sub
 
Private Sub OptBtnMoreThanOneFile_Click()
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    
    FilePath = Me.txtLink.Text
    
    'Loop Through Files
    If Trim(FilePath) <> "" Then
        With UserFormOpenFiles
            .lstFileList.Visible = True
            .OptBtnFileRead.Value = False
            .OptBtnFileWrite.Value = False
            .Frame2.Enabled = True
        End With

        Me.lstFileList.Clear
        
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oFolder = oFSO.GetFolder(FilePath)
        
        For Each oFile In oFolder.Files
            
            With Me.lstFileList
                If Right(oFile.Name, 5) = ".xlsx" Or Right(oFile.Name, 5) = ".xlsm" Or Right(oFile.Name, 4) = ".xls" Then
                    .AddItem oFile.Name
                End If
            End With
        Next oFile
    End If
End Sub


Private Sub cmdOpen_Click()
    FilePath = Me.txtLink.Text
    
    If Me.OptBtnFile.Value = False And Me.OptBtnFolder.Value = False And Me.OptBtnMoreThanOneFile.Value = False Then
        MsgBox "One of File or Folder Options must be selected."
        Exit Sub
   ElseIf (Me.OptBtnFile.Value = True Or Me.OptBtnMoreThanOneFile.Value = True) And Me.OptBtnFileRead.Value = False And Me.OptBtnFileWrite = False Then
        MsgBox "File read-write option must be selected."
        Exit Sub
    ElseIf Trim(Me.txtLink.Text) = "" Then
        MsgBox "Link field is empty"
        Exit Sub
    
    'More Than One File Open
    ElseIf Me.OptBtnMoreThanOneFile.Value = True Then
        For i = 0 To Me.lstFileList.ListCount - 1
            If Me.lstFileList.Selected(i) Then
                FileLink = FilePath & "\" & Me.lstFileList.List(i)

                If Me.OptBtnFileRead.Value = True Then
                    Workbooks.Open FileLink, , True
                ElseIf Me.OptBtnFileWrite.Value = True Then
                    Workbooks.Open FileLink, , False
                Else
                    MsgBox "Unknown option :("
                End If
                
                Unload UserFormOpenFiles
            End If
        Next

    'One File Open
    ElseIf Me.OptBtnFile.Value = True Then
        If Me.OptBtnFileRead.Value = True Then
            Workbooks.Open FilePath, , True
        Else
            Workbooks.Open FilePath, , False
        End If

        Unload UserFormOpenFiles

    'Folder Open
    ElseIf Me.OptBtnFolder.Value = True Then
        Shell "C:\WINDOWS\explorer.exe """ & FilePath & "", vbNormalFocus
    Else
        MsgBox "Unknow option :("
    End If
End Sub

