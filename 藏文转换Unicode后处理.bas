Attribute VB_Name = "����ת��Unicode����"
Sub ת������Unicode֮��Ĵ���()
'
' ת������Unicode֮��Ĵ���
'����
    Selection.Find.Execute findtext:=ChrW(3857), replacewith:=ChrW(3853), Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    '���ߴ���
    Selection.Find.Execute findtext:=ChrW(3853), replacewith:=ChrW(3853) & " ", Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    
    Selection.Find.Execute findtext:="   ", replacewith:=" ", Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    Selection.Find.Execute findtext:="  ", replacewith:=" ", Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    
    Selection.Find.Execute findtext:=ChrW(3851) & " ", replacewith:=ChrW(3851), Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    
    Selection.Find.Execute findtext:=ChrW(3853) & " " & ChrW(3853) & " ", replacewith:=ChrW(3853) & " " & ChrW(3853), Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    Selection.Find.Execute findtext:=ChrW(3853) & " " & ChrW(3853) & " ", replacewith:=ChrW(3853) & " " & ChrW(3853), Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    Selection.Find.Execute findtext:=ChrW(3853) & " " & ChrW(3853) & ChrW(3853) & " " & ChrW(3853), replacewith:=ChrW(3854) & " " & ChrW(3854), Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    Selection.Find.Execute findtext:=" " & ChrW(3853) & " ", replacewith:=" " & ChrW(3853), Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    Selection.Find.Execute findtext:=" " & ChrW(3853) & " ", replacewith:=" " & ChrW(3853), Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    
    Selection.Find.Execute findtext:=ChrW(3906) & ChrW(3853), replacewith:=ChrW(3906) & " " & ChrW(3853), Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    Selection.Find.Execute findtext:=ChrW(3906) & " " & ChrW(3853) & " ", replacewith:=ChrW(3906) & " " & ChrW(3853), Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    Selection.Find.Execute findtext:=ChrW(3860), replacewith:=ChrW(3860) & " ", Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    Selection.Find.Execute findtext:=ChrW(3860) & " " & ChrW(3860) & " ", replacewith:=ChrW(3860) & " " & ChrW(3860), Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    
    Selection.Find.Execute findtext:=ChrW(3906) & " " & ChrW(3853) & ChrW(3853) & ChrW(3853), replacewith:=ChrW(3906) & ChrW(3853) & " " & ChrW(3853) & ChrW(3853), Format:=True, Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    MsgBox ("Done!")

'    With Selection.Find
'        .ClearFormatting
'        .Text = ""
'        .Font.Name = ""
'        With .Replacement
'            .ClearFormatting
'            .Font.Name = ""
'        End With
'        .Execute findtext:="", replacewith:="", Format:=True, Replace:=wdReplaceAll
'    End With

End Sub

