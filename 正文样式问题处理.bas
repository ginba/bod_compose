Attribute VB_Name = "正文样式问题处理"
Sub GaiSBTluanma()
Dim numberA
numberA = InputBox("大约要改多少字符？请输入数字")
    For n = 1 To numberA
        With Selection
            .MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.Font.Name = "Calibri" Then Selection.Font.Name = "dedris-a"
            .MoveRight
        End With
    Next n
    MsgBox ("Done!")
End Sub
Sub 正文样式处理()
    With ActiveDocument.Styles("正文").Font
        .NameAscii = "Calibri"
        .NameOther = "Calibri"
        .Name = "Calibri"
        .Color = wdColorAutomatic
    End With
    MsgBox ("Done!")
End Sub
Sub C2A()
Attribute C2A.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.宏1"
'
' 宏1 宏
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Font.Name = "Calibri"
        .Replacement.Font.Name = "Dedris-a"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Font.Name = "+西文正文"
        .Replacement.Font.Name = "Dedris-a"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "-"
        .Replacement.Font.Name = "Dedris-a"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub
Sub C2Vowa()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Font.Name = "Calibri"
        .Replacement.Font.Name = "Dedris-vowa"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Font.Name = "Times New Roman"
        .Replacement.Font.Name = "Dedris-vowa"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Font.Name = "Calibri (西文正文)"
        .Replacement.Font.Name = "Dedris-vowa"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "-"
        .Replacement.Font.Name = "Dedris-vowa"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub
