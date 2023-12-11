Attribute VB_Name = "数式変換モジュール"
Const strAI As String = "C:\AI\"
Const strTex As String = "eq.tex"
Const strBat As String = "MakeEPS.bat"
Const strSvg As String = "eq.svg"

Sub 数式をTimes系で変換()
    Call ProcessConvertMath("newtx.tex")
End Sub

Private Sub ProcessConvertMath(ByVal strTexTemplate As String)
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim tmpShape As Shape
        For Each tmpShape In ActiveWindow.Selection.ShapeRange
            If tmpShape.TextFrame.HasText Then
                Call CreateNewTexFile(strTexTemplate, tmpShape.TextFrame.TextRange.Text)
                Call CreateSvgFile
                Call InsertSvgFile(tmpShape)
            Else
                If tmpShape.AlternativeText <> "" Then
                    With ActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeRectangle, tmpShape.Left, tmpShape.Top, 10, 10)
                        .TextFrame.TextRange.Text = tmpShape.AlternativeText
                        .TextFrame.WordWrap = msoFalse
                        .TextFrame.AutoSize = ppAutoSizeShapeToFitText
                    End With
                End If
            End If
        Next tmpShape
    End If
End Sub

Private Sub CreateNewTexFile(ByVal strTexTemplate As String, ByVal strTexText As String)
    Dim strNewTexText As String
    Dim buf As String
    
    '元のファイルを読み込み
    Open strAI & strTexTemplate For Input As #1
    Do Until EOF(1)
        Line Input #1, buf
        If strNewTexText <> "" Then
            strNewTexText = strNewTexText & vbNewLine
        End If
        If buf = "<r>" Then
            buf = strTexText
        End If
        strNewTexText = strNewTexText & buf
    Loop
    Close #1
    
    '新しいファイルに書き込み
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .WriteText strNewTexText, 1
        .SaveToFile strAI & strTex, 2
        .Close
    End With
End Sub

Private Sub CreateSvgFile()
    With CreateObject("Wscript.Shell")
        Call .Run(strAI & strBat, WaitOnReturn:=True)
    End With
End Sub

Private Sub InsertSvgFile(ByRef tmpShape As Shape)
    If CreateObject("Scripting.FileSystemObject").FileExists(strAI & strSvg) Then
        Dim tmpPicture As Shape
        With ActiveWindow.Selection.SlideRange.Shapes.AddPicture(strAI & strSvg, msoFalse, msoTrue, 0, 0)
            .Left = tmpShape.Left + tmpShape.Width / 2 - .Width / 2
            .Top = tmpShape.Top + tmpShape.Height / 2 - .Height / 2
            .AlternativeText = tmpShape.TextFrame.TextRange.Text
            tmpShape.Delete
        End With
    End If
End Sub

