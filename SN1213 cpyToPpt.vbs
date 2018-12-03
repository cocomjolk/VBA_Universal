Sub copyPasteCommentTable()

    Dim pptApp      As PowerPoint.Application
    Dim pptPres     As PowerPoint.Presentation
    Dim pptSlide    As PowerPoint.Slide
    Dim pptShape    As PowerPoint.Shape
    Dim pptLayout   As PowerPoint.CustomLayout
    Dim toCopy      As Range
    Dim ws          As Worksheet
    Dim i           As Long 'slide counter
    Dim hiValue     As Integer


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
    'pptApp.Visible = True

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

    Set pptSlide = pptPres.Slides.AddSlide(i + 1, pptLayout)
    'Set pptSlide = pptPres.Slides.Add(1, ppLayout)

    hiValue = 1

    If hiValue = 1 Then
        Sheets("Template (1)").Select
        ActiveSheet.UsedRange.Select
        'Range("A1:K4").Select
        Selection.Copy
        pptSlide.Shapes.Paste
        Set pptShape = pptSlide.Shapes(1)
        With pptShape
            .LockAspectRatio = msoTrue
            .Left = 10
            .Top = -80
            .Width = 925
            .Height = 200
        End With

     ElseIf hiValue = 2 Then
        Sheets("Template (2)").Select
        ActiveSheet.UsedRange.Select
        'Range("A1:J5").Select
        Selection.Copy
        pptSlide.Shapes.Paste
        Set pptShape = pptSlide.Shapes(1)
        With pptShape
            .LockAspectRatio = msoTrue
            .Left = 10
            .Top = -50
            .Width = 920
            .Height = 300
        End With

     ElseIf hiValue = 3 Then
        Sheets("Template (3)").Select
        ActiveSheet.UsedRange.Select
        'Range("A1:J6").Select
        Selection.Copy
        pptSlide.Shapes.Paste
        Set pptShape = pptSlide.Shapes(1)
        With pptShape
            .LockAspectRatio = msoTrue
            .Left = 10
            .Top = -50
            .Width = 920
            .Height = 400
        End With

     ElseIf hiValue = 4 Then
        Sheets("Template (4)").Select
        ActiveSheet.UsedRange.Select
        'Range("A1:J7").Select
        Selection.Copy
        pptSlide.Shapes.Paste
        Set pptShape = pptSlide.Shapes(1)
        With pptShape
            .LockAspectRatio = msoTrue
            .Left = 10
            .Top = -40
            .Width = 920
            .Height = 500
        End With

     ElseIf hiValue = 5 Then
        Sheets("Template (5)").Select
        ActiveSheet.UsedRange.Select
        'Range("A1:K7").Select
        Selection.Copy
        pptSlide.Shapes.Paste
        Set pptShape = pptSlide.Shapes(1)
        With pptShape
            .LockAspectRatio = msoTrue
            .Left = 10
            .Top = -40
            .Width = 920
            .Height = 500
        End With

    End If
    Set pptApp = Nothing

End Sub
