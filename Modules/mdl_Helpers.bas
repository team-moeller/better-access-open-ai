Attribute VB_Name = "mdl_Helpers"
'###########################################################################################
'# Copyright (c) 2023 Thomas Moeller                                                       #
'# MIT License  => https://github.com/team-moeller/better-access-open-ai/blob/main/LICENSE #
'# Version 0.91.04  published: 06.08.2023                                                  #
'###########################################################################################

#If VBA7 Then
    Private Declare PtrSafe Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
#Else
    Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
#End If

Public Sub PrepareAndExportModules(Optional ByVal TagVersion As Boolean = True)
' This method is intended for the power user who wants to update the version numbers in the code headers

    'Declarations
    Dim Version As String
    Dim CodeLine As String
    Dim vbc As Object
    
    MakeSureDirectoryPathExists CurrentProject.Path & "\Modules\"
    Version = DLast("V_Number", "tbl_VersionHistory")
    CodeLine = "'# Version " & Version & "  published: " & Format$(Date, "dd.mm.yyyy") & "                                                  #"
    
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        If vbc.Type = 1 Or vbc.Type = 2 Then
            If TagVersion Then
                Application.VBE.ActiveVBProject.VBComponents(vbc.Name).CodeModule.InsertLines 4, CodeLine
                Application.VBE.ActiveVBProject.VBComponents(vbc.Name).CodeModule.DeleteLines 5, 1
            End If
            Application.VBE.ActiveVBProject.VBComponents(vbc.Name).Export CurrentProject.Path & "\Modules\" & vbc.Name & IIf(vbc.Type = 2, ".cls", ".bas")
        End If
    Next
    Application.VBE.ActiveVBProject.VBComponents("Form_frm_Demo").CodeModule.InsertLines 6, "    Me.txtAPI_Key = ""INSERT YOUR OPEN AI API KEY HERE"""
    Application.VBE.ActiveVBProject.VBComponents("Form_frm_Demo").CodeModule.DeleteLines 7, 1
    
    Application.DoCmd.RunCommand (acCmdCompileAndSaveAllModules)
    
    MsgBox "Export done", vbInformation, "Better Access Open-AI"

End Sub
