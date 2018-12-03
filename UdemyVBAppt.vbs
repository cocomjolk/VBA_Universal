Sub CreatePptSlide()
    Dim pptApp      As PowerPoint.Application
    Dim pptPres     As PowerPoint.Presentation
    Dim pptSlide    As PowerPoint.Slide
    Dim pptShape    As PowerPoint.Shape
    Dim pptLayout   As PowerPoint.CustomLayout
    Dim toCopy      As Range
    Dim wS          As Worksheet
    Dim i           As Long 'slide counter

    'this command ignores error.
    On Error Resume Next
    'this command assumes ppt is open. If not error will be ignored
    Set pptApp = GetObject(, "PowerPoint.Application")
    'this code restarts error handling.
    On Error GoTo 0
    'open a new ppt session if not open
    If pptApp Is Nothing Then
        Set pptApp = New PowerPoint.Application
    End If
    'get presentation
    'setting if condition to check if a presentation exists..
    'by counting for presentations
    If pptApp.Presentations.Count = 0 Then
        ' if true then add/create a new presentation
        Set pptPres = pptApp.Presentations.Add
        i = 0
    Else
        Set pptPres = pptApp.ActivePresentation
        'counts the slides of active presentation
        i = pptPres.Slides.Count
    End If
    'setting pptLayout as layout 7, as determined my numerical order in ppt app when creating a new slide.
    Set pptLayout = pptPres.SlideMaster.CustomLayouts(7)
    'going to loop through all worksheets that exist after worksheet(7)
    For Each wS In ThisWorkbook.Worksheets
        If wS.Index > 7 Then
            Set toCopy = wS.UsedRange
            Set pptSlide = pptPres.Slides.AddSlide(i + 1, pptLayout)
            toCopy.CopyPicture xlScreen, xlPicture
            pptSlide.Shapes.Paste
            Set pptShape = pptSlide.Shapes(1)
            With pptShape
                .LockAspectRatio = msoTrue
                .Left = 100
                .Top = 200
                .Width = 500
            End With
            i = i + 1
        End If
    Next wS
    'MsgBox "your slide was created"

    Set pptApp = Nothing

End Sub
