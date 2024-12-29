Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Excel.Range, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    Dim myF As Variant
    Cancel = True
    
    If Target.Columns.Count <> 23 Or Target.Rows.Count <> 19 Then Exit Sub
    
    myF = Application.GetOpenFilename _
    ("jpg jpeg bmp tif png gif,*.jpg;*.jpeg;*.bmp;*.tif;*.png;*.gif", , "画像の選択", , False)
    If myF = False Then Exit Sub
    
    With ActiveSheet.Shapes.AddPicture(Filename:=myF, LinkToFile:=False, _
        SaveWithDocument:=True, Left:=Target.Left, Top:=Target.Top, _
        Width:=-1, Height:=-1)
        
        .LockAspectRatio = True  '縦横比率を維持する
        
        '画像の元のアスペクト比を計算
        Dim imageRatio As Double
        imageRatio = .Width / .Height
        
        'セルのアスペクト比を計算
        Dim cellRatio As Double
        cellRatio = Target.Width / Target.Height
        
        '画像をセルに合わせてリサイズ
        If imageRatio > cellRatio Then
            '画像の方が横長の場合
            .Width = Target.Width
            .Height = Target.Width / imageRatio
        Else
            '画像の方が縦長の場合
            .Height = Target.Height
            .Width = Target.Height * imageRatio
        End If
        
        '中央に配置
        .Top = Target.Top + (Target.Height - .Height) / 2
        .Left = Target.Left + (Target.Width - .Width) / 2
    End With
    
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
    Cancel = True
End Sub
