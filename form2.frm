Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set gallery.Picture = LoadPicture("C:\Users\SHREYA-lappy\Desktop\outline\galopen.jpg")
    
    If Me.WindowState <> vbMinimized Then
        'gallery.Picture.Width = Me.Width
        'gallery.Picture.Height = Me.Height
    End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set gallery.Picture = LoadPicture("C:\Users\SHREYA-lappy\Desktop\outline\galsave.jpg")
    
    If Me.WindowState <> vbMinimized Then
        'gallery.Picture.Width = Me.Width
        'gallery.Picture.Height = Me.Height
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set gallery.Picture = LoadPicture("C:\Users\SHREYA-lappy\Desktop\outline\PhotoGalleryHeader2.jpg")
    'gallery.AutoRedraw=True
    'me.PaintPicture me.Picture,0,0,me.Width,me.Height
    
    If Me.WindowState <> vbMinimized Then
        'gallery.Picture.Width = Me.Width
        'gallery.Picture.Height = Me.Height
    End If
End Sub
Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set gallery.Picture = LoadPicture("C:\Users\SHREYA-lappy\Desktop\outline\galprint.jpg")
    
    If Me.WindowState <> vbMinimized Then
        'gallery.Picture.Width = Me.Width
        'gallery.Picture.Height = Me.Height
    End If
End Sub
