Attribute VB_Name = "cut1"
Sub 解析取代的()
Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.content
End If

 mystr = area.find.Execute(findtext:="^#^#.^t", ReplaceWith:="答案：", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="^#.^t", ReplaceWith:="答案：", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:=" ", ReplaceWith:="", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="答案：", ReplaceWith:="答案：" & Chr(13), Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="答案：", ReplaceWith:="答案：*", Replace:=wdReplaceAll)
  Call 解析用ABCD換行
  
 End Sub
Sub 解析用ABCD換行()
Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.content
End If


 mystr = area.find.Execute(findtext:="：(", ReplaceWith:="xx", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:=Chr(13) & "(A)", ReplaceWith:="(A)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:=Chr(13) & "(B)", ReplaceWith:="(B)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:=Chr(13) & "(C)", ReplaceWith:="(C)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:=Chr(13) & "(D)", ReplaceWith:="(D)", Replace:=wdReplaceAll)
 
 mystr = area.find.Execute(findtext:="(A)", ReplaceWith:="#(A)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(B)", ReplaceWith:="#(B)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(C)", ReplaceWith:="#(C)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(D)", ReplaceWith:="#(D)", Replace:=wdReplaceAll)

 
 mystr = area.find.Execute(findtext:="#(", ReplaceWith:=Chr(13) & "(", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="xx", ReplaceWith:="：(", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(E)", ReplaceWith:=Chr(13) & "(E)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:=")" & Chr(13) & "(", ReplaceWith:=")(", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="選" & Chr(13) & "(", ReplaceWith:="選(", Replace:=wdReplaceAll)
 Dim count As Integer
    count = 0
    For Each para In ActiveDocument.Paragraphs

        If InStr(para.Range.Text, "答案：") > 0 Then
            count = count + 1
        End If
    
    Next para
    
    MsgBox "答案：有" & count & " 個。"
End Sub
Sub 解析一次用()
   Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.content
End If
mystr = area.find.Execute(findtext:="答案：", ReplaceWith:="", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="　", ReplaceWith:=Chr(13) & "　", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="　", ReplaceWith:="", Replace:=wdReplaceAll)
Documents("解析.docx").Activate
Selection.HomeKey wdStory
 charmoved = Selection.EndOf(Unit:=wdParagraph, Extend:=wdExtend)
 Selection.MoveStart Unit:=wdCharacter, count:=3
 Selection.MoveEnd Unit:=wdCharacter, count:=-1
 Selection.Delete
 Documents("答案1.docx").Activate
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
            ' 將計數器加1
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
 Documents("解析.docx").Activate
 Call 找到答案
Documents("答案1.docx").Activate
Next
Documents.Open fileName:="C:\Users\Pinyan\Desktop\解析.docx"
Call 解析改成標楷含貼完取代
Call 解析選項取代成全型
End Sub
Sub 找到答案()
Selection.find.Execute findtext = "答案：", foward = True
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

Sub 解析改成標楷含貼完取代()
   ActiveDocument.content.Select
    Selection.Font.Name = "標楷體"
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    Selection.ParagraphFormat.LineSpacing = 18
 Selection.ParagraphFormat.CharacterUnitLeftIndent = 2
    Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.content
End If
mystr = area.find.Execute(findtext:="答案：.", ReplaceWith:="答案：", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="答案：*.", ReplaceWith:="答案：", Replace:=wdReplaceAll)
End Sub
Sub 題組解析選項改全形()
  Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.content
End If
mystr = area.find.Execute(findtext:="(1)", ReplaceWith:="(１)", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="(2)", ReplaceWith:="(２)", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="(3)", ReplaceWith:="(３)", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="(4)", ReplaceWith:="(４)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(A)", ReplaceWith:="(Ａ)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(B)", ReplaceWith:="(Ｂ)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="(C)", ReplaceWith:="(Ｃ)", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:="(D)", ReplaceWith:="(Ｄ)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="(E)", ReplaceWith:="(Ｅ)", Replace:=wdReplaceAll)

End Sub
Sub 解析選項格式改動()
 Dim oPara As Paragraph
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(Ａ)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(Ｂ)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(Ｃ)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(Ｄ)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
   For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "(Ｅ)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.Text, "【語譯】") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("002")
    End If
  Next
Call 答案標號
End Sub

