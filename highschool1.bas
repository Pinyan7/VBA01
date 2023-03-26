Attribute VB_Name = "highschool1"
Sub 插入D句號()
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
        myRange.InsertAfter "。"
        myRange.Collapse wdCollapseEnd
    End If
Loop
Call 插入E句號
Call 取代五四三的
Call 將題目自動編號
Call 更改選項的樣式
End Sub

Sub 插入E句號()
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
        myRange.InsertAfter "。"
        myRange.Collapse wdCollapseEnd
Loop
End Sub

Sub 取代五四三的()
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
     mystr = area.find.Execute(findtext:="★", ReplaceWith:="", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="「」", ReplaceWith:="「　」", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="：(A)", ReplaceWith:="：" & Chr(13) & "(A)", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="。」。", ReplaceWith:="。」", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="！」。", ReplaceWith:="！」", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="？」。", ReplaceWith:="？」", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="(^#)()", ReplaceWith:="", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="(共0分,每題0分)", ReplaceWith:="", Replace:=wdReplaceAll)
       mystr = area.find.Execute(findtext:="※", ReplaceWith:="【篇章】", Replace:=wdReplaceAll)
       ActiveDocument.Content.Select
   Selection.Style = ActiveDocument.Styles("00")
       
End Sub
Sub 解析取代的()
Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
End If

 mystr = area.find.Execute(findtext:="^#^#.^t", ReplaceWith:="答案：", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="^#.^t", ReplaceWith:="答案：", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:=" ", ReplaceWith:="", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="答案：", ReplaceWith:="答案：" & Chr(13), Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="答案：", ReplaceWith:="答案：*", Replace:=wdReplaceAll)
  Call 解析用ABCD換行
  
 End Sub
 Sub 解析選項取代成全型()
 Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
End If
  mystr = area.find.Execute(findtext:="A", ReplaceWith:="(Ａ)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="B", ReplaceWith:="(Ｂ)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="C", ReplaceWith:="(Ｃ)", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:="D", ReplaceWith:="(Ｄ)", Replace:=wdReplaceAll)
    mystr = area.find.Execute(findtext:="E", ReplaceWith:="(Ｅ)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(A)", ReplaceWith:="(Ａ)", Replace:=wdReplaceAll)
 mystr = area.find.Execute(findtext:="(B)", ReplaceWith:="(Ｂ)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="(C)", ReplaceWith:="(Ｃ)", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:="(D)", ReplaceWith:="(Ｄ)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="(E)", ReplaceWith:="(Ｅ)", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="((", ReplaceWith:="(", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="))", ReplaceWith:=")", Replace:=wdReplaceAll)
 End Sub
  Sub 題組解析用取代()
   Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
End If
Call RangeReplace
mystr = area.find.Execute(findtext:="^#.", ReplaceWith:="", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="答案：", ReplaceWith:="【篇章】答案：", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="解析：", ReplaceWith:="", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:=")　", ReplaceWith:=")；", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:="^t", ReplaceWith:="", Replace:=wdReplaceAll)
  mystr = area.find.Execute(findtext:=" ", ReplaceWith:="", Replace:=wdReplaceAll)
    mystr = area.find.Execute(findtext:=")" & Chr(13) & "(A", ReplaceWith:=")(A", Replace:=wdReplaceAll)
   mystr = area.find.Execute(findtext:=")" & Chr(13) & "(B", ReplaceWith:=")(B", Replace:=wdReplaceAll)
    mystr = area.find.Execute(findtext:=")" & Chr(13) & "(C", ReplaceWith:=")(C", Replace:=wdReplaceAll)
     mystr = area.find.Execute(findtext:=")" & Chr(13) & "(D", ReplaceWith:=")(D", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:=")" & Chr(13) & "(E", ReplaceWith:=")(E", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="(共0分,每題0分)", ReplaceWith:="", Replace:=wdReplaceAll)
      mystr = area.find.Execute(findtext:="選" & Chr(13) & "(", ReplaceWith:="選(", Replace:=wdReplaceAll)
      Call 題組用解析格式改動
  End Sub

Sub 將題目自動編號()
  Dim oPara As paragraph
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(A)") > 0 Then
     oPara.Previous.Range.Style = ActiveDocument.Styles("1")
    End If
  Next
End Sub

Sub 更改選項的樣式()
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
Sub 解析一次用()
   Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
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
            ' 將計數器加1
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
 Documents("解析.docx").Activate
 Call 找到答案
Documents("答案1.docx").Activate
Next
Documents.Open FileName:="C:\Users\User\Desktop\國文\高中段複卷\高一段複(三民)\解析.docx"
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
   ActiveDocument.Content.Select
    Selection.Font.Name = "標楷體"
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    Selection.ParagraphFormat.LineSpacing = 18
 Selection.ParagraphFormat.CharacterUnitLeftIndent = 2
    Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
End If
mystr = area.find.Execute(findtext:="答案：.", ReplaceWith:="答案：", Replace:=wdReplaceAll)
mystr = area.find.Execute(findtext:="答案：*.", ReplaceWith:="答案：", Replace:=wdReplaceAll)
End Sub
Sub 答案標號()
Dim doc As Word.Document
    Dim Range As Word.Range
    Dim count As Integer
  Dim findtext As String
   Set doc = ActiveDocument
  findtext = "答案："
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
            ' 將計數器加1
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
 .Execute findtext:="答案："
  End With
  charmoved = Selection.EndOf(Unit:=wdParagraph, Extend:=wdExtend)
  Selection.ClearFormatting
Selection.ParagraphFormat.LineUnitBefore = 0.5
Selection.Font.Name = "新細明體"
Selection.Font.Size = 14
Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    Selection.ParagraphFormat.LineSpacing = 25
 Selection.Range.ListFormat.ApplyNumberDefault
     With Selection.Range.ListFormat
        .ListTemplate.ListLevels(1).Font.Name = "標楷體"
        .ListTemplate.ListLevels(1).Font.Bold = True
    End With
     Selection.EndOf Unit:=wdWord, Extend:=wdMove
 Next
End Sub

Sub 解析選項格式改動()
 Dim oPara As paragraph
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(Ａ)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(Ｂ)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(Ｃ)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
    For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(Ｄ)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
   For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "(Ｅ)") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("001")
    End If
  Next
  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "【語譯】") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("002")
    End If
  Next
Call 答案標號
End Sub

Sub 解析用ABCD換行()
Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
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

        If InStr(para.Range.text, "答案：") > 0 Then
            count = count + 1
        End If
    
    Next para
    
    MsgBox "答案：有" & count & " 個。"
End Sub
Sub 題組用解析格式改動()
Dim oPara As paragraph
ActiveDocument.Content.Style = ActiveDocument.Styles("一括號一選項ABCD")
For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "二、") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("0")
    End If
  Next
     For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "【語譯】") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("題組解析用語譯")
    End If
  Next
 Set regEx = CreateObject("VBScript.RegExp")
 With regEx
    .Pattern = "\([1-9]\)"
    .Global = True
End With
For Each oPara In ActiveDocument.Paragraphs
    If regEx.test(oPara.Range.text) Then
        oPara.Range.Style = ActiveDocument.Styles("數字括號")
    End If
Next oPara
With regEx
    .Pattern = "\([1-9]\)\([A-Z]\)"
    .Global = True
End With
For Each oPara In ActiveDocument.Paragraphs
    If regEx.test(oPara.Range.text) Then
        oPara.Range.Style = ActiveDocument.Styles("一括號一選項")
    End If
Next oPara

  For Each oPara In ActiveDocument.Paragraphs
    If InStr(oPara.Range.text, "【篇章") > 0 Then
      oPara.Range.Style = ActiveDocument.Styles("題組解析答案")
    End If
  Next
  Call 題組解析選項改全形
End Sub
Sub 題組解析選項改全形()
  Dim mystr As String
Set area = Selection.Range '
If area.Start = area.End Then
 Set area = ActiveDocument.Content
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
Sub InsertPageBreakAfterParagraphD()
    ' 定義變量
    Dim para As paragraph
    Dim targetPara As paragraph
    
    ' 設置目標段落
    For Each para In ActiveDocument.Paragraphs
        If para.Range.text = "(D)" Then ' 使用 vbCr 以包含段落結尾符號
            Set targetPara = para
            Exit For
        End If
    Next para
    
    ' 在目標段落後插入分頁符號
    If Not targetPara Is Nothing Then
        targetPara.Range.InsertBreak Type:=wdPageBreak
    Else
        MsgBox "未找到目標段落。"
    End If
End Sub
