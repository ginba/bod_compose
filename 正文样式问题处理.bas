Attribute VB_Name = "������ʽ���⴦��"
Sub GaiSBTluanma()
Dim numberA
numberA = InputBox("��ԼҪ�Ķ����ַ�������������")
    For n = 1 To numberA
        With Selection
            .MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.Font.Name = "Calibri" Then Selection.Font.Name = "dedris-a"
            .MoveRight
        End With
    Next n
    MsgBox ("Done!")
End Sub
Sub ������ʽ����()
    With ActiveDocument.Styles("����").Font
        .NameAscii = "Calibri"
        .NameOther = "Calibri"
        .Name = "Calibri"
        .Color = wdColorAutomatic
    End With
    MsgBox ("Done!")
End Sub
Sub C2A()
Attribute C2A.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.��1"
'
' ��1 ��
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
        .Font.Name = "+��������"
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
        .Font.Name = "Calibri (��������)"
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
