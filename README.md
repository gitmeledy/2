Sub ExtractImagesToJPG()
    Dim doc As Document
    Dim shp As InlineShape
    Dim imgPath As String
    Dim counter As Integer
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim pptShape As Object

    imgPath = "C:\макросы\картинки\" ' Укажите нужную папку
    
    If Dir(imgPath, vbDirectory) = "" Then MkDir imgPath

    Set doc = ActiveDocument
    counter = 1
    
    ' Запускаем PowerPoint
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application")
    End If
    On Error GoTo 0
    
    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Add

    For Each shp In doc.InlineShapes
        If shp.Type = wdInlineShapePicture Then
            shp.Range.Copy

            ' Добавляем слайд и вставляем картинку
            Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, 1)
            pptSlide.Shapes(1).Delete ' Удаляем заголовок
            pptSlide.Shapes(1).Delete ' Удаляем подзаголовок
            
            Set pptShape = pptSlide.Shapes.PasteSpecial(DataType:=2)(1) ' 2 = ppPasteEnhancedMetafile
            
            ' Настраиваем размер слайда под изображение
            pptSlide.FollowMasterBackground = msoFalse
            pptShape.LockAspectRatio = msoTrue
            pptShape.Left = 0
            pptShape.Top = 0
            pptShape.ScaleHeight 1, msoTrue
            pptShape.ScaleWidth 1, msoTrue

            ' Сохраняем слайд как изображение
            pptSlide.Export imgPath & "Image" & counter & ".jpg", "JPG"
            counter = counter + 1
        End If
    Next shp

    pptPres.Close
    pptApp.Quit
    Set pptApp = Nothing

    MsgBox "Изображения успешно извлечены!", vbInformation
End Sub
# 2
