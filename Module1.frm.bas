Attribute VB_Name = "Module1"
Sub ExportSlidesToImageAndText()
    Dim pptSlide As Slide
    Dim pptShape As Shape
    Dim exportPath As String
    Dim fileName As String
    Dim txtContent As String
    Dim fso As Object
    Dim txtFile As Object
    
    'set the storage path(default is the "Export_Result" folder under the presentation folder)
    exportPath = ActivePresentation.Path & "\Export_Result\"
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath
    
    ' establish a text file object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(exportPath & "All_Slides_Text.txt", True, True)

    ' iterate through each slide
    For Each pptSlide In ActivePresentation.Slides
        ' 1. export slide to image files(png format)
        fileName = "Slide_" & pptSlide.SlideIndex & ".png"
        pptSlide.Export exportPath & fileName, "PNG"
        
        ' 2. extract the text from the slides
        txtContent = "--- Slide " & pptSlide.SlideIndex & " ---" & vbCrLf
        For Each pptShape In pptSlide.Shapes
            If pptShape.HasTextFrame Then
                If pptShape.TextFrame.HasText Then
                    txtContent = txtContent & pptShape.TextFrame.TextRange.Text & vbCrLf
                End If
            End If
        Next pptShape
        
        ' write the text into a txt file
        txtFile.WriteLine txtContent & vbCrLf
    Next pptSlide

    txtFile.Close
    MsgBox "Export completed ! The file is saved at: " & vbCrLf & exportPath, vbInformation
End Sub

