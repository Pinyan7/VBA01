Attribute VB_Name = "ReplaceText"
Sub RangeReplace()
Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
End If


 mystr = area.Find.Execute(findtext:="¡G(", replacewith:="xx", Replace:=wdReplaceAll)
 mystr = area.Find.Execute(findtext:=Chr(13) & "(A)", replacewith:="(A)", Replace:=wdReplaceAll)
 mystr = area.Find.Execute(findtext:=Chr(13) & "(B)", replacewith:="(B)", Replace:=wdReplaceAll)
 mystr = area.Find.Execute(findtext:=Chr(13) & "(C)", replacewith:="(C)", Replace:=wdReplaceAll)
 mystr = area.Find.Execute(findtext:=Chr(13) & "(D)", replacewith:="(D)", Replace:=wdReplaceAll)
 
 mystr = area.Find.Execute(findtext:="(A)", replacewith:="#(A)", Replace:=wdReplaceAll)
 mystr = area.Find.Execute(findtext:="(B)", replacewith:="#(B)", Replace:=wdReplaceAll)
 mystr = area.Find.Execute(findtext:="(C)", replacewith:="#(C)", Replace:=wdReplaceAll)
 mystr = area.Find.Execute(findtext:="(D)", replacewith:="#(D)", Replace:=wdReplaceAll)
 
 mystr = area.Find.Execute(findtext:="#(", replacewith:=Chr(13) & "(", Replace:=wdReplaceAll)
 mystr = area.Find.Execute(findtext:="xx", replacewith:="¡G(", Replace:=wdReplaceAll)
 
End Sub

