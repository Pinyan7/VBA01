Attribute VB_Name = "cut1"
Sub �ѪR���N��()
Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.content
End If

 mystr = area.find.Execute(findtext:="^#^#.^t", ReplaceWith:="���סG", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="^#.^t", ReplaceWith:="���סG", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:=" ", ReplaceWith:="", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="���סG", ReplaceWith:="���סG" & Chr(13), Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="���סG", ReplaceWith:="���סG*", Replace:=wdReplaceAll)
  Call �ѪR��ABCD����
  
 End Sub
Sub �ѪR��ABCD����()
Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.content
End If


 mystr = area.find.Execute(findtext:="�G(", ReplaceWith:="xx", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:=Chr(13) & "(A)", ReplaceWith:="(A)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:=Chr(13) & "(B)", ReplaceWith:="(B)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:=Chr(13) & "(C)", ReplaceWith:="(C)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:=Chr(13) & "(D)", ReplaceWith:="(D)", Replace:=wdReplaceAll)
 
 mystr = area.find.Execute(findtext:="(A)", ReplaceWith:="#(A)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(B)", ReplaceWith:="#(B)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(C)", ReplaceWith:="#(C)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(D)", ReplaceWith:="#(D)", Replace:=wdReplaceAll)

 
 mystr = area.find.Execute(findtext:="#(", ReplaceWith:=Chr(13) & "(", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="xx", ReplaceWith:="�G(", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(E)", ReplaceWith:=Chr(13) & "(E)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:=")" & Chr(13) & "(", ReplaceWith:=")(", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="��" & Chr(13) & "(", ReplaceWith:="��(", Replace:=wdReplaceAll)
 Dim count As Integer
    count = 0
    For Each para In ActiveDocument.Paragraphs

        If InStr(para.Range.Text, "���סG") > 0 Then
            count = count + 1
        End If
    
    Next para
    
    MsgBox "���סG��" & count & " �ӡC"
End Sub
Sub �ѪR�@����()
   Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.content
End If
mystr = area.find.Execute(findtext:="���סG", ReplaceWith:="", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="�@", ReplaceWith:=Chr(13) & "�@", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="�@", ReplaceWith:="", Replace:=wdReplaceAll)
Documents("�ѪR.docx").Activate
Selection.HomeKey wdStory
 charmoved = Selection.EndOf(Unit:=wdParagraph, Extend:=wdExtend)
 Selection.MoveStart Unit:=wdCharacter, count:=3
 Selection.MoveEnd Unit:=wdCharacter, count:=-1
 Selection.Delete
 Documents("����1.docx").Activate
 Selection.HomeKey wdStory
Dim doc As Word.Document
    Dim Range As Word.Range
    Dim count As Integer
  Dim findtext As String
   Set doc = ActiveDocument
  findtext = "."
  count = 0
  Set Range = doc.content
   With Range.find
        .Text = findtext
        .Forward = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        While .Execute
            ' �N�p�ƾ��[1
            count = count + 1
        Wend
    End With
  For I = 1 To count
  With Selection.find
 .ClearFormatting
 .Wrap = wdFindContinue
 .Text = findtext
 .Execute
  End With
 charmoved = Selection.EndOf(Unit:=wdParagraph, Extend:=wdExtend)
 Selection.MoveEnd Unit:=wdCharacter, count:=-1
 Selection.CUT
 Documents("�ѪR.docx").Activate
 Call ��쵪��
Documents("����1.docx").Activate
Next
Documents.Open fileName:="C:\Users\Pinyan\Desktop\�ѪR.docx"
Call �ѪR�令�з��t�K�����N
Call �ѪR�ﶵ���N������
End Sub
Sub ��쵪��()
Selection.find.Execute findtext = "���סG", foward = True
Selection.EndOf Unit:=wdWord, Extend:=wdMove
 Selection.Paste
 With Selection.find
 .Forward = True
 .ClearFormatting
 .MatchWholeWord = True
 .MatchCase = False
 .Wrap = wdFindContinue
 .Execute findtext:="*"
  End With
End Sub

Sub �ѪR�令�з��t�K�����N()
   ActiveDocument.content.Select
    Selection.Font.Name = "�з���"
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    Selection.ParagraphFormat.LineSpacing = 18
 Selection.ParagraphFormat.CharacterUnitLeftIndent = 2
    Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.content
End If
mystr = area.find.Execute(findtext:="���סG.", ReplaceWith:="���סG", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="���סG*.", ReplaceWith:="���סG", Replace:=wdReplaceAll)
End Sub
Sub �D�ոѪR�ﶵ�����()
  Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.content
End If
mystr = area.find.Execute(findtext:="(1)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="(2)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="(3)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="(4)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(A)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(B)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="(C)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:="(D)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="(E)", ReplaceWith:="(��)", Replace:=wdReplaceAll)

End Sub
Sub �ѪR�ﶵ�榡���()
 Dim oPara As Paragraph
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(��)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(��)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(��)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(��)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
   For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(��)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "�i�yĶ�j") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("002")
    End If
  Next
Call ���׼и�
End Sub

