Option Explicit

' Main: Batch process all presentations in the folder
Sub BatchExportAllPresentations()
    Dim folderPath As String
    Dim fileName As String
    Dim pres As Presentation
    Dim processedCount As Long
    Dim errorCount As Long
    
    ' Initialize counters
    processedCount = 0
    errorCount = 0
    
    On Error GoTo ErrorHandler
    
    ' Specify folder path (using the current presentation's folder)
    folderPath = ActivePresentation.Path
    If folderPath = "" Then
        MsgBox "Please save the current presentation first to determine the working folder location.", vbExclamation
        Exit Sub
    End If
    
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' ========== Processing pptx files ==========
    fileName = Dir(folderPath & "*.pptx")
    Do While fileName <> ""
        If fileName <> "." And fileName <> ".." Then
            If Len(Dir(folderPath & fileName)) > 0 Then
                On Error Resume Next
                Set pres = Nothing
                
                ' Check if this is the currently active presentation
                If folderPath & fileName = ActivePresentation.FullName Then
                    Debug.Print "Processing current presentation: " & fileName
                    Set pres = ActivePresentation
                    Call ExportSlidesToImageAndText_Enhanced_ForPres(pres)
                    processedCount = processedCount + 1
                Else
                    Debug.Print "Opening: " & folderPath & fileName
                    Set pres = Presentations.Open(folderPath & fileName, _
                        ReadOnly:=msoTrue, Untitled:=msoFalse, WithWindow:=msoFalse)
                    
                    If Err.Number = 0 Then
                        Call ExportSlidesToImageAndText_Enhanced_ForPres(pres)
                        pres.Close
                        processedCount = processedCount + 1
                    Else
                        Debug.Print "Could not open: " & fileName & " - " & Err.Description
                        errorCount = errorCount + 1
                        Err.Clear
                    End If
                End If
                
                On Error GoTo ErrorHandler
            End If
        End If
        fileName = Dir
    Loop
    
    ' ========== Processing pptm files ==========
    fileName = Dir(folderPath & "*.pptm")
    Do While fileName <> ""
        If fileName <> "." And fileName <> ".." Then
            If Len(Dir(folderPath & fileName)) > 0 Then
                On Error Resume Next
                Set pres = Nothing
                
                If folderPath & fileName = ActivePresentation.FullName Then
                    Debug.Print "Processing current presentation: " & fileName
                    Set pres = ActivePresentation
                    Call ExportSlidesToImageAndText_Enhanced_ForPres(pres)
                    processedCount = processedCount + 1
                Else
                    Debug.Print "Opening: " & folderPath & fileName
                    Set pres = Presentations.Open(folderPath & fileName, _
                        ReadOnly:=msoTrue, Untitled:=msoFalse, WithWindow:=msoFalse)
                    
                    If Err.Number = 0 Then
                        Call ExportSlidesToImageAndText_Enhanced_ForPres(pres)
                        pres.Close
                        processedCount = processedCount + 1
                    Else
                        Debug.Print "Could not open: " & fileName & " - " & Err.Description
                        errorCount = errorCount + 1
                        Err.Clear
                    End If
                End If
                
                On Error GoTo ErrorHandler
            End If
        End If
        fileName = Dir
    Loop
    
    ' ========== Processing legacy ppt files ==========
    fileName = Dir(folderPath & "*.ppt")
    Do While fileName <> ""
        If fileName <> "." And fileName <> ".." Then
            If Len(Dir(folderPath & fileName)) > 0 Then
                On Error Resume Next
                Set pres = Nothing
                
                If folderPath & fileName = ActivePresentation.FullName Then
                    Debug.Print "Processing current presentation: " & fileName
                    Set pres = ActivePresentation
                    Call ExportSlidesToImageAndText_Enhanced_ForPres(pres)
                    processedCount = processedCount + 1
                Else
                    Debug.Print "Opening: " & folderPath & fileName
                    Set pres = Presentations.Open(folderPath & fileName, _
                        ReadOnly:=msoTrue, Untitled:=msoFalse, WithWindow:=msoFalse)
                    
                    If Err.Number = 0 Then
                        Call ExportSlidesToImageAndText_Enhanced_ForPres(pres)
                        pres.Close
                        processedCount = processedCount + 1
                    Else
                        Debug.Print "Could not open: " & fileName & " - " & Err.Description
                        errorCount = errorCount + 1
                        Err.Clear
                    End If
                End If
                
                On Error GoTo ErrorHandler
            End If
        End If
        fileName = Dir
    Loop
    
    ' Display completion message
    If errorCount = 0 Then
        MsgBox "All presentations have been exported successfully!" & vbCrLf & _
               "Files processed: " & processedCount, vbInformation
    Else
        MsgBox "Batch processing completed with some errors." & vbCrLf & _
               "Success: " & processedCount & " file(s)" & vbCrLf & _
               "Failed: " & errorCount & " file(s)" & vbCrLf & _
               "Please check the Immediate Window (Ctrl+G) for detailed error information.", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An unexpected error occurred during batch processing: " & Err.Number & " - " & Err.Description, vbCritical
    If Not pres Is Nothing Then
        On Error Resume Next
        pres.Close
    End If
End Sub


' Subroutine: Export slides, images, and text from a single presentation
Sub ExportSlidesToImageAndText_Enhanced_ForPres(pres As Presentation)
    Dim pptSlide As Slide
    Dim pptShape As Shape
    Dim exportPath As String
    Dim baseFileName As String
    Dim slidePngName As String
    Dim picFileName As String
    Dim txtFileName As String
    Dim fso As Object
    Dim txtAll As Object
    Dim txtSlide As Object
    Dim picCounter As Long
    Dim slideIndexStr As String
    Dim dotPos As Long
    Dim slideTextContent As String
    Dim shapeText As String
    
    On Error GoTo LocalErrorHandler
    
    ' Get filename (without extension)
    dotPos = InStrRev(pres.Name, ".")
    If dotPos > 0 Then
        baseFileName = Left(pres.Name, dotPos - 1)
    Else
        baseFileName = pres.Name
    End If
    
    ' Clean illegal characters from filename (to avoid folder name errors)
    baseFileName = CleanFileName(baseFileName)
    
    ' Create export folder
    exportPath = pres.Path & "\Export_Result\" & baseFileName & "\"
    CreateFolderRecursive exportPath
    
    Dim pathWholeSlide As String, pathSlideText As String, pathSlideImage As String
    pathWholeSlide = exportPath & "whole_slide\"
    pathSlideText = exportPath & "slide_text\"
    pathSlideImage = exportPath & "slide_image\"
    
    CreateFolderRecursive pathWholeSlide
    CreateFolderRecursive pathSlideText
    CreateFolderRecursive pathSlideImage
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtAll = fso.CreateTextFile(pathSlideText & "All_Slides_Text.txt", True, True)
    
    ' Process each slide
    For Each pptSlide In pres.Slides
        slideIndexStr = Format(pptSlide.SlideIndex, "00")
        
        ' Export entire slide as PNG
        slidePngName = baseFileName & "_Slide_" & slideIndexStr & ".png"
        pptSlide.Export pathWholeSlide & slidePngName, "PNG"
        
        ' Export images
        picCounter = 0
        For Each pptShape In pptSlide.Shapes
            If pptShape.Type = msoPicture Or pptShape.Type = msoLinkedPicture Or pptShape.Type = msoOLEControlObject Then
                On Error Resume Next
                picCounter = picCounter + 1
                picFileName = baseFileName & "_p" & slideIndexStr & "_" & Format(picCounter, "00") & ".png"
                
                ' Export using original dimensions to maintain quality
                pptShape.Export pathSlideImage & picFileName, ppShapeFormatPNG
                
                If Err.Number <> 0 Then
                    Debug.Print "  Could not export image: Slide " & slideIndexStr & " Shape " & pptShape.Name
                    Err.Clear
                End If
                On Error GoTo LocalErrorHandler
            End If
        Next pptShape
        
        ' Collect text content
        slideTextContent = "--- Slide " & pptSlide.SlideIndex & " ---" & vbCrLf & vbCrLf
        
        For Each pptShape In pptSlide.Shapes
            If pptShape.HasTextFrame Then
                If pptShape.TextFrame.HasText Then
                    shapeText = Trim(pptShape.TextFrame.TextRange.Text)
                    If Len(shapeText) > 0 Then
                        slideTextContent = slideTextContent & _
                                           "Shape ID: " & pptShape.Id & "  |  Name: " & pptShape.Name & vbCrLf & _
                                           "----------------------------------------" & vbCrLf & _
                                           shapeText & vbCrLf & vbCrLf
                    End If
                End If
            End If
            
            ' Process text within grouped shapes
            If pptShape.Type = msoGroup Then
                Call ProcessGroupShapes(pptShape, slideTextContent)
            End If
        Next pptShape
        
        ' Write individual slide text file
        txtFileName = baseFileName & "_p" & slideIndexStr & ".txt"
        Set txtSlide = fso.CreateTextFile(pathSlideText & txtFileName, True, True)
        txtSlide.Write slideTextContent
        txtSlide.Close
        Set txtSlide = Nothing
        
        ' Write to consolidated text file
        txtAll.WriteLine slideTextContent & String(80, "=") & vbCrLf
    Next pptSlide
    
    txtAll.Close
    Set txtAll = Nothing
    Set fso = Nothing
    
    Debug.Print "Export completed: " & pres.Name
    Exit Sub
    
LocalErrorHandler:
    MsgBox "Error occurred while exporting presentation (" & pres.Name & "): " & Err.Description, vbExclamation
    If Not txtSlide Is Nothing Then txtSlide.Close
    If Not txtAll Is Nothing Then txtAll.Close
    Set fso = Nothing
End Sub


' Helper: Recursively process shapes within groups
Private Sub ProcessGroupShapes(grpShape As Shape, ByRef textContent As String)
    Dim subShape As Shape
    Dim shapeText As String
    
    On Error Resume Next
    
    For Each subShape In grpShape.GroupItems
        If subShape.HasTextFrame Then
            If subShape.TextFrame.HasText Then
                shapeText = Trim(subShape.TextFrame.TextRange.Text)
                If Len(shapeText) > 0 Then
                    textContent = textContent & _
                                  "[Grouped] Shape ID: " & subShape.Id & "  |  Name: " & subShape.Name & vbCrLf & _
                                  "----------------------------------------" & vbCrLf & _
                                  shapeText & vbCrLf & vbCrLf
                End If
            End If
        End If
        
        ' Recursively process nested groups
        If subShape.Type = msoGroup Then
            Call ProcessGroupShapes(subShape, textContent)
        End If
    Next subShape
End Sub


' Helper: Remove illegal characters from filename
Private Function CleanFileName(originalName As String) As String
    Dim invalidChars As Variant
    Dim i As Long
    Dim result As String
    
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    result = originalName
    
    For i = LBound(invalidChars) To UBound(invalidChars)
        result = Replace(result, invalidChars(i), "_")
    Next i
    
    CleanFileName = result
End Function


' Helper: Recursively create folder structure
Private Sub CreateFolderRecursive(folderPath As String)
    Dim fso As Object
    Dim parentPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Exit if folder already exists
    If fso.FolderExists(folderPath) Then Exit Sub
    
    ' Create parent folder first (recursive)
    parentPath = Left(folderPath, InStrRev(folderPath, "\", Len(folderPath) - 1) - 1)
    If Len(parentPath) > 3 Then ' Skip drive letters (e.g., C:\)
        If Not fso.FolderExists(parentPath) Then
            CreateFolderRecursive parentPath
        End If
    End If
    
    ' Create this level folder
    MkDir folderPath
    Set fso = Nothing
End Sub

