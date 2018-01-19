Attribute VB_Name = "GitImport"
Sub pImport()
    Dim ModulePath As String
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim cmpComponent As VBIDE.VBComponent
    
    If ThisWorkbook.VBProject.Protection = 1 Then
        MsgBox "Workbook is protected against the importing codes!!" & _
        "go to the Tools>Macros>Security(in Excel), click on the Trusted Publishers tab and check the Trust access to the Visual Basic Project setting"
        Exit Sub
    End If
    
    ModulePath = ThisWorkbook.Path & "\"
    
    'Deleting Existing Modules before importing
    For Each cmpComponent In ActiveWorkbook.VBProject.VBComponents
        If cmpComponent.Type = vbext_ct_StdModule And cmpComponent.Name <> "GitImport" Then
            ThisWorkbook.VBProject.VBComponents.Remove cmpComponent
        End If
    Next cmpComponent
    
    ' Import the VBA code required
    If objFSO.GetFolder(ModulePath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    Else
        For Each objFile In objFSO.GetFolder(ModulePath).Files
            If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
                (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
                (objFSO.GetExtensionName(objFile.Name) = "bas") Then
                ThisWorkbook.VBProject.VBComponents.Import objFile.Path
            End If
        Next objFile
    End If
    ActiveWorkbook.VBProject.VBComponents.Import ModulePath
    
    ' Save the workbook
    ActiveWorkbook.Save
    MsgBox "Import Successful!"
End Sub
Sub pExport()
    Dim bExport As Boolean
    Dim ModulePath As String
    Dim cmpComponent As VBIDE.VBComponent
    Dim szFileName As String
    ModulePath = ThisWorkbook.Path
    
    Kill ModulePath & "\*.bas"
    
    If ThisWorkbook.VBProject.Protection = 1 Then
        MsgBox "Workbook is protected against the importing codes!!" & _
        "go to the Tools>Macros>Security(in Excel), click on the Trusted Publishers tab and check the Trust access to the Visual Basic Project setting"
        Exit Sub
    End If
    For Each cmpComponent In ThisWorkbook.VBProject.VBComponents
        bExport = True
        szFileName = cmpComponent.Name
        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export ModulePath & "\" & szFileName
        End If
    Next cmpComponent
    
End Sub


