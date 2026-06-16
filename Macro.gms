Sub FitPageToSelection_Perfect_X7()
    Dim s As ShapeRange
    Set s = ActiveSelectionRange
    
    ' N?u chua ch?n d?i tu?ng nào thì thoát
    If s.count = 0 Then
        MsgBox "Vui long chon doi tuong truoc!", vbExclamation, "Thong bao"
        Exit Sub
    End If
    
    ' Ð?t don v? do chu?n là Millimeter
    ActiveDocument.Unit = cdrMillimeter
    
    ' 1. Ghi nh? v? trí g?c (tâm X và tâm Y) c?a d?i tu?ng tru?c khi d?ch chuy?n
    Dim luuX As Double, luuY As Double
    luuX = s.PositionX
    luuY = s.PositionY
    
    ' 2. Ð?i kích thu?c trang gi?y b?ng dúng kích thu?c d?i tu?ng
    ActivePage.SetSize s.SizeWidth, s.SizeHeight
    
    ' 3. Ðua d?i tu?ng vào gi?a trang d? l?y m?c t?a d? m?i cho trang gi?y
    s.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
    
    ' 4. Tính toán kho?ng l?ch và d?ch chuy?n trang gi?y d?n v? trí cu c?a d?i tu?ng
    ' (Th?c ch?t là d?ch chuy?n toàn b? th? gi?i t?a d? quay v? v? trí ban d?u)
    ActiveDocument.BeginCommandGroup "Fit Page"
    Dim deltaX As Double, deltaY As Double
    deltaX = luuX - s.PositionX
    deltaY = luuY - s.PositionY
    
    ' Tr? d?i tu?ng v? v? trí cu so v?i màn hình, kéo theo trang gi?y di cùng
    s.Move deltaX, deltaY
    ActiveDocument.EndCommandGroup
    
End Sub
Sub SelectAllBitmaps()
    Dim s As Shape
    Dim srAsBitmaps As New ShapeRange
    
    ' Quét qua toàn b? các d?i tu?ng có trong trang hi?n t?i
    For Each s In ActivePage.Shapes
        ' N?u d?i tu?ng dó thu?c ki?u Bitmap (hình ?nh)
        If s.Type = cdrBitmapShape Then
            srAsBitmaps.Add s
        End If
    Next s
    
    ' N?u tìm th?y hình ?nh thì t? d?ng ch?n (bôi den) t?t c? chúng
    If srAsBitmaps.count > 0 Then
        srAsBitmaps.CreateSelection
        MsgBox "Da tim va chon " & srAsBitmaps.count & " file hinh anh trong trang!", vbInformation, "Thong bao"
    Else
        MsgBox "Khong tim thay file hinh anh (Bitmap) nao trong trang nay!", vbExclamation, "Thong bao"
    End If
End Sub
Sub FindLowResolutionBitmaps()
    Dim s As Shape
    Dim lowResShapes As New ShapeRange
    Dim dpiX As Long, dpiY As Long
    
    ' Quét qua t?t c? d?i tu?ng trong trang hi?n t?i
    For Each s In ActivePage.Shapes
        ' N?u d?i tu?ng là hình ?nh (Bitmap)
        If s.Type = cdrBitmapShape Then
            ' L?y thông s? dpi th?c t? c?a ?nh
            dpiX = s.Bitmap.ResolutionX
            dpiY = s.Bitmap.ResolutionY
            
            ' N?u dpi nh? hon 300 thì gom vào danh sách
            If dpiX < 300 Or dpiY < 300 Then
                lowResShapes.Add s
            End If
        End If
    Next s
    
    ' X? lý k?t qu?
    If lowResShapes.count > 0 Then
        ' T? d?ng ch?n t?t c? các ?nh b? thi?u dpi
        lowResShapes.CreateSelection
        MsgBox "Da tim va chon " & lowResShapes.count & " hinh anh co do phan giai duoi 300 dpi!", vbInformation, "Thong bao"
    Else
        MsgBox "Tuyet voi! Tat ca hinh anh trong trang deu da dat tu 300 dpi tro len.", vbInformation, "Thong bao"
    End If
End Sub
Sub FindLowRes_HeavyFile()
    Dim allBitmaps As ShapeRange
    Dim lowResShapes As New ShapeRange
    Dim s As Shape
    
    ' T?i uu hóa h? th?ng d? không b? do/treo màn hình khi x? lý file n?ng
    Optimization = True
    On Error Resume Next
    
    ' L?nh siêu t?c: Tìm T?T C? các Bitmap trong toàn b? trang (K? c? n?m trong Group)
    Set allBitmaps = ActivePage.FindShapes(Type:=cdrBitmapShape)
    
    ' N?u không có ?nh nào thì thoát luôn
    If allBitmaps.count = 0 Then
        Optimization = False
        MsgBox "Khong tim thay file hinh anh nao!", vbInformation, "Thong bao"
        Exit Sub
    End If
    
    ' Quét danh sách ?nh tìm du?c d? l?c ra ?nh du?i 300 dpi
    For Each s In allBitmaps
        If s.Bitmap.ResolutionX < 300 Or s.Bitmap.ResolutionY < 300 Then
            lowResShapes.Add s
        End If
    Next s
    
    ' Tr? l?i giao di?n Corel
    Optimization = False
    ActiveWindow.Refresh
    
    ' Hi?n th? k?t qu? và t? d?ng ch?n (Select)
    If lowResShapes.count > 0 Then
        lowResShapes.CreateSelection
        MsgBox "Da tim va chon nhanh " & lowResShapes.count & " anh duoi 300 dpi giua hang nghin doi tuong!", vbInformation, "Thong bao"
    Else
        MsgBox "Tat ca hinh anh trong file deu da dat tu 300 dpi tro len!", vbInformation, "Thong bao"
    End If
End Sub
