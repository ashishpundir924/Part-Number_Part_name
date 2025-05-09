Sub RenamePartNamesToFileNames()

    Dim invApp As Inventor.Application
    Set invApp = ThisApplication

    Dim doc As Document
    Dim partDoc As PartDocument
    Dim fileName As String

    On Error GoTo ErrorHandler

    For Each doc In invApp.Documents
        If doc.DocumentType = kPartDocumentObject Then
            Set partDoc = doc

            ' Get file name without extension or path
            fileName = Left(partDoc.FullFileName, InStrRev(partDoc.FullFileName, ".") - 1)
            fileName = Mid(fileName, InStrRev(fileName, "\") + 1)

            ' Update iProperty: Part Number
            partDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = fileName
            
            ' Optionally save
            ' partDoc.Save
        End If
    Next doc

    MsgBox "All open part files updated!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical

End Sub
