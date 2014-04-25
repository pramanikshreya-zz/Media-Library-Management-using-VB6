Private Sub Form_resize()

    Set Image1.Picture = LoadPicture("C:\Users\SHREYA-lappy\Desktop\outline\instagram41.jpg")
    
    If Me.WindowState <> vbMinimized Then
        Image1.Width = Me.Width
        Image1.Height = Me.Height
    End If

End Sub

