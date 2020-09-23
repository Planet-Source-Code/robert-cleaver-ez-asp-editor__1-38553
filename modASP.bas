Attribute VB_Name = "Module1"
Option Explicit
Global Saved As Boolean
Global OldPath As String
Function SetStatus(Status As String)
    frmMain.statusbar.Caption = " :: " & Status
End Function

Function OpenNew()
    Dim Filex
    Dim Buffer$
    Dim FilePath$

    Filex = FreeFile
    frmMain.fileDLG.ShowOpen
    FilePath$ = frmMain.fileDLG.FileName
    If FilePath$ = ("") Then Exit Function
    SetStatus ("Opening  " & frmMain.fileDLG.FileTitle)
    Open FilePath$ For Binary Access Read As Filex
        Buffer$ = Space(LOF(Filex))
        Get #Filex, , Buffer$
    Close #Filex
    
    frmMain.txtpad.Text = Buffer$
    Buffer$ = ("")
    FilePath$ = ("")
    Saved = True
    OldPath = ("")
    SetStatus ("Opened")
End Function

Function SaveASP()
Dim Filex
    Filex = FreeFile
    If OldPath = ("") Then Exit Function
    SetStatus ("Saving " & OldPath)
    Open OldPath For Output As Filex
        Print #Filex, frmMain.txtpad.Text
    Close Filex
    SetStatus ("Saved..")
End Function

Function LoadSnippets()
    
End Function
    

Function SaveAs()
    Dim Filex
    Dim Buffer$
    Dim FilePath$
    Filex = FreeFile
    frmMain.fileDLG.ShowOpen
    FilePath$ = frmMain.fileDLG.FileName
    If FilePath = ("") Then Exit Function
    SetStatus ("Saving " & frmMain.fileDLG.FileTitle)
    Open FilePath$ For Output As Filex
        Print #Filex, frmMain.txtpad.Text
    Close Filex
    Saved = True
    Buffer$ = ("")
    FilePath$ = ("")
    OldPath = FilePath
    SetStatus ("Saved..")
End Function
