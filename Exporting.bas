Attribute VB_Name = "Exporting"
    Sub ExportAllModules()
        Dim vbProj As Object
        Dim vbComp As Object
        Dim sPath As String
        Dim sFileName As String

        ' Set the path where you want to export the modules
        ' You can change ThisWorkbook.Path to a specific folder path if desired
        sPath = "C:\Users\Muz\Documents\ExportedOldModules2\"

        ' Create the folder if it doesn't exist
        If Dir(sPath, vbDirectory) = "" Then
            MkDir sPath
        End If

        Set vbProj = ThisWorkbook.VBProject

        For Each vbComp In vbProj.VBComponents
            ' Only export standard modules, class modules, and userforms
            Select Case vbComp.Type
                Case 1 ' vbext_ct_StdModule (Standard Module)
                    sFileName = sPath & vbComp.Name & ".bas"
                Case 2 ' vbext_ct_ClassModule (Class Module)
                    sFileName = sPath & vbComp.Name & ".cls"
                Case 3 ' vbext_ct_MSForm (UserForm)
                    sFileName = sPath & vbComp.Name & ".frm"
                Case Else
                    ' Skip other component types (e.g., ThisWorkbook, Sheet objects)
                    GoTo NextComponent
            End Select

            vbComp.Export sFileName
            Debug.Print "Exported: " & sFileName

NextComponent:
        Next vbComp

        MsgBox "All modules exported to: " & sPath, vbInformation
    End Sub
