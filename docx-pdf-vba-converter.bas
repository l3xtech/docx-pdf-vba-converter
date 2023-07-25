Sub ConvertDocxToPDF()
    Dim folderPath As String
    Dim docxFile As String
    Dim pdfFile As String
    Dim doc As Document

    ' Change this to the path of the directory that contains the .docx files
    folderPath = "C:\path\to\your\files\"

    ' Make sure the path ends with a backslash
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath + "\"
    End If

    docxFile = Dir(folderPath & "*.docx", vbNormal)

    While docxFile <> ""
        Set doc = Documents.Open(folderPath & docxFile)
        
        ' Construct the .pdf file name
        pdfFile = folderPath & Replace(docxFile, ".docx", ".pdf")

        ' Save as PDF (wdFormatPDF = 17)
        doc.SaveAs pdfFile, FileFormat:=17
        
        doc.Close SaveChanges:=wdDoNotSaveChanges

        ' Next .docx file
        docxFile = Dir()
    Wend
End Sub

