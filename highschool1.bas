Attribute VB_Name = "highschool1"
Sub ���JD�y��()
    Dim myRange As Range
    Set myRange = ActiveDocument.Content
    Call RangeReplace
    
    With myRange.find
        .ClearFormatting
        .text = "(D)"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
 Do While myRange.find.Execute
    ' Check if next paragraph contains "(E)"
    If InStr(myRange.Next(wdParagraph).text, "(E)") = 0 Then
        ' Insert period after paragraph if next paragraph doesn't contain "(E)"
        myRange.MoveEnd wdParagraph, 1
        myRange.MoveEnd wdCharacter, -1
        myRange.Collapse wdCollapseEnd
        myRange.InsertAfter "�C"
        myRange.Collapse wdCollapseEnd
    End If
Loop
Call ���JE�y��
Call ���N���|�T��
Call �N�D�ئ۰ʽs��
Call ���ﶵ���˦�
End Sub

Sub ���JE�y��()
    Dim myRange As Range
    Set myRange = ActiveDocument.Content
    
    With myRange.find
        .ClearFormatting
        .text = "(E)"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
 Do While myRange.find.Execute
        myRange.MoveEnd wdParagraph, 1
        myRange.MoveEnd wdCharacter, -1
        myRange.Collapse wdCollapseEnd
        myRange.InsertAfter "�C"
        myRange.Collapse wdCollapseEnd
Loop
End Sub

Sub ���N���|�T��()
Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
End If
 mystr = area.find.Execute(findtext:="^#^#.", ReplaceWith:="", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="^#.", ReplaceWith:="", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:="^t", ReplaceWith:="", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:="()", ReplaceWith:="", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:=" ", ReplaceWith:="", Replace:=wdReplaceAll)
     mystr = area.find.Execute(findtext:="��", ReplaceWith:="", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="�u�v", ReplaceWith:="�u�@�v", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="�G(A)", ReplaceWith:="�G" & Chr(13) & "(A)", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="�C�v�C", ReplaceWith:="�C�v", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="�I�v�C", ReplaceWith:="�I�v", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="�H�v�C", ReplaceWith:="�H�v", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="(^#)()", ReplaceWith:="", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="(�@0��,�C�D0��)", ReplaceWith:="", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="��", ReplaceWith:="�i�g���j", Replace:=wdReplaceAll)
       ActiveDocument.Content.Select
   Selection.Style = ActiveDocument.Styles("00")
       
End Sub
Sub �ѪR���N��()
Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
End If

 mystr = area.find.Execute(findtext:="^#^#.^t", ReplaceWith:="���סG", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="^#.^t", ReplaceWith:="���סG", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:=" ", ReplaceWith:="", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="���סG", ReplaceWith:="���סG" & Chr(13), Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="���סG", ReplaceWith:="���סG*", Replace:=wdReplaceAll)
  Call �ѪR��ABCD����
  
 End Sub
 Sub �ѪR�ﶵ���N������()
 Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
End If
  mystr = area.find.Execute(findtext:="A", ReplaceWith:="(��)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="B", ReplaceWith:="(��)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="C", ReplaceWith:="(��)", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:="D", ReplaceWith:="(��)", Replace:=wdReplaceAll)
    mystr = area.find.Execute(findtext:="E", ReplaceWith:="(��)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(A)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(B)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="(C)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:="(D)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="(E)", ReplaceWith:="(��)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="((", ReplaceWith:="(", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="))", ReplaceWith:=")", Replace:=wdReplaceAll)
 End Sub
  Sub �D�ոѪR�Ψ��N()
   Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
End If
Call RangeReplace
mystr = area.find.Execute(findtext:="^#.", ReplaceWith:="", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="���סG", ReplaceWith:="�i�g���j���סG", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="�ѪR�G", ReplaceWith:="", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:=")�@", ReplaceWith:=")�F", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="^t", ReplaceWith:="", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:=" ", ReplaceWith:="", Replace:=wdReplaceAll)
    mystr = area.find.Execute(findtext:=")" & Chr(13) & "(A", ReplaceWith:=")(A", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:=")" & Chr(13) & "(B", ReplaceWith:=")(B", Replace:=wdReplaceAll)
    mystr = area.find.Execute(findtext:=")" & Chr(13) & "(C", ReplaceWith:=")(C", Replace:=wdReplaceAll)
     mystr = area.find.Execute(findtext:=")" & Chr(13) & "(D", ReplaceWith:=")(D", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:=")" & Chr(13) & "(E", ReplaceWith:=")(E", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="(�@0��,�C�D0��)", ReplaceWith:="", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="��" & Chr(13) & "(", ReplaceWith:="��(", Replace:=wdReplaceAll)
      Call �D�եθѪR�榡���
  End Sub

Sub �N�D�ئ۰ʽs��()
  Dim oPara As paragraph
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(A)") > 0 Then
     oPara.Previous.Range.Style = ActiveDocument.Styles("1")
    End If
  Next
End Sub

Sub ���ﶵ���˦�()
  Dim oPara As paragraph
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(A)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("2")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(B)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("2")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(C)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("2")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(D)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("2")
    End If
  Next
   For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(E)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("2")
    End If
  Next
End Sub
Sub �ѪR�@����()
   Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
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
  Set Range = doc.Content
   With Range.find
        .text = findtext
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
  For i = 1 To count
  With Selection.find
 .ClearFormatting
 .Wrap = wdFindContinue
 .text = findtext
 .Execute
  End With
 charmoved = Selection.EndOf(Unit:=wdParagraph, Extend:=wdExtend)
 Selection.MoveEnd Unit:=wdCharacter, count:=-1
 Selection.CUT
 Documents("�ѪR.docx").Activate
 Call ��쵪��
Documents("����1.docx").Activate
Next
Documents.Open FileName:="C:\Users\User\Desktop\���\�����q�ƨ�\���@�q��(�T��)\�ѪR.docx"
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
   ActiveDocument.Content.Select
    Selection.Font.Name = "�з���"
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    Selection.ParagraphFormat.LineSpacing = 18
 Selection.ParagraphFormat.CharacterUnitLeftIndent = 2
    Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
End If
mystr = area.find.Execute(findtext:="���סG.", ReplaceWith:="���סG", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="���סG*.", ReplaceWith:="���סG", Replace:=wdReplaceAll)
End Sub
Sub ���׼и�()
Dim doc As Word.Document
    Dim Range As Word.Range
    Dim count As Integer
  Dim findtext As String
   Set doc = ActiveDocument
  findtext = "���סG"
  count = 0
  Set Range = doc.Content
   With Range.find
        .text = findtext
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
    Selection.HomeKey wdStory
For i = 1 To count
With Selection.find
 .Forward = True
 .ClearFormatting
 .MatchWholeWord = True
 .MatchCase = False
 .Wrap = wdFindContinue
 .Execute findtext:="���סG"
  End With
  charmoved = Selection.EndOf(Unit:=wdParagraph, Extend:=wdExtend)
  Selection.ClearFormatting
Selection.ParagraphFormat.LineUnitBefore = 0.5
Selection.Font.Name = "�s�ө���"
Selection.Font.Size = 14
Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    Selection.ParagraphFormat.LineSpacing = 25
 Selection.Range.ListFormat.ApplyNumberDefault
     With Selection.Range.ListFormat
        .ListTemplate.ListLevels(1).Font.Name = "�з���"
        .ListTemplate.ListLevels(1).Font.Bold = True
    End With
     Selection.EndOf Unit:=wdWord, Extend:=wdMove
 Next
End Sub

Sub �ѪR�ﶵ�榡���()
 Dim oPara As paragraph
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(��)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(��)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(��)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(��)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
   For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(��)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "�i�yĶ�j") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("002")
    End If
  Next
Call ���׼и�
End Sub

Sub �ѪR��ABCD����()
Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
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

        If InStr(para.Range.text, "���סG") > 0 Then
            count = count + 1
        End If
    
    Next para
    
    MsgBox "���סG��" & count & " �ӡC"
End Sub
Sub �D�եθѪR�榡���()
Dim oPara As paragraph
ActiveDocument.Content.Style = ActiveDocument.Styles("�@�A���@�ﶵABCD")
For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "�G�B") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("0")
    End If
  Next
     For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "�i�yĶ�j") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("�D�ոѪR�λyĶ")
    End If
  Next
 Set regEx = CreateObject("VBScript.RegExp")
 With regEx
    .Pattern = "\([1-9]\)"
    .Global = True
End With
For Each oPara In ActiveDocument.Paragraphs
    If regEx.test(oPara.Range.text) Then
        oPara.Range.Style = ActiveDocument.Styles("�Ʀr�A��")
    End If
Next oPara
With regEx
    .Pattern = "\([1-9]\)\([A-Z]\)"
    .Global = True
End With
For Each oPara In ActiveDocument.Paragraphs
    If regEx.test(oPara.Range.text) Then
        oPara.Range.Style = ActiveDocument.Styles("�@�A���@�ﶵ")
    End If
Next oPara

  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "�i�g��") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("�D�ոѪR����")
    End If
  Next
  Call �D�ոѪR�ﶵ�����
End Sub
Sub �D�ոѪR�ﶵ�����()
  Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
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
Sub InsertPageBreakAfterParagraphD()
    ' �w�q�ܶq
    Dim para As paragraph
    Dim targetPara As paragraph
    
    ' �]�m�ؼЬq��
    For Each para In ActiveDocument.Paragraphs
        If para.Range.text = "(D)" Then ' �ϥ� vbCr �H�]�t�q�������Ÿ�
            Set targetPara = para
            Exit For
        End If
    Next para
    
    ' �b�ؼЬq���ᴡ�J�����Ÿ�
    If Not targetPara Is Nothing Then
        targetPara.Range.InsertBreak Type:=wdPageBreak
    Else
        MsgBox "�����ؼЬq���C"
    End If
End Sub
