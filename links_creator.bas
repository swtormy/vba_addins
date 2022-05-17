Sub create_link()
On Error Resume Next
Dim sld As Slide
Dim shp As Shape
Dim hshp As Hyperlink
Dim SlideInfo As String

SlideInfo = InputBox("Ââåäèòå íàçâàíèÿ ññûëîê è íîìåð ñëàéäà (<Ñëàéä íàçíà÷åíèÿ>_<¹ ñëàéäà íàçíà÷åíèÿ>_<Îáðàòíàÿ ññûëêà>, íàïðèìåð: Öåëè_12_Ãëàâíàÿ)")

If SlideInfo <> "" Then
If CInt(Split(SlideInfo, "_")(1)) <= Application.ActivePresentation.Slides.Count Then
mySlide = Application.ActiveWindow.View.Slide.SlideNumber

Application.ActiveWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Left:=610, Top:=4, Width:=160, Height:=27).Name = Split(SlideInfo, "_")(0)
Application.ActiveWindow.View.Slide.Shapes.Range(Split(SlideInfo, "_")(0)).Line.Weight = 1.5
Application.ActiveWindow.View.Slide.Shapes.Range(Split(SlideInfo, "_")(0)).Line.ForeColor.RGB = RGB(250, 200, 0)
Application.ActiveWindow.View.Slide.Shapes.Range(Split(SlideInfo, "_")(0)).Fill.ForeColor.RGB = RGB(255, 255, 255)
Application.ActiveWindow.View.Slide.Shapes.Range(Split(SlideInfo, "_")(0)).Fill.BackColor.RGB = RGB(250, 200, 0)
Application.ActiveWindow.View.Slide.Shapes.Range(Split(SlideInfo, "_")(0)).TextFrame.TextRange.Text = Split(SlideInfo, "_")(0)
Application.ActiveWindow.View.Slide.Shapes.Range(Split(SlideInfo, "_")(0)).TextFrame.TextRange.Font.Color.RGB = RGB(250, 200, 0)
Application.ActiveWindow.View.Slide.Shapes.Range(Split(SlideInfo, "_")(0)).TextFrame.TextRange.Font.Size = 14
Application.ActiveWindow.View.Slide.Shapes.Range(Split(SlideInfo, "_")(0)).TextFrame.TextRange.Font.Name = "Arial"
Application.ActiveWindow.View.Slide.Shapes.Range(Split(SlideInfo, "_")(0)).TextFrame.TextRange.Font.Underline = msoTrue

Application.ActiveWindow.View.Slide.Shapes.Range(Split(SlideInfo, "_")(0)).ActionSettings(ppMouseClick).Action = ppActionHyperlink
Application.ActiveWindow.View.Slide.Shapes.Range(Split(SlideInfo, "_")(0)).ActionSettings(ppMouseClick).Hyperlink.SubAddress = ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).SlideNumber

ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.AddShape(Type:=msoShapeRectangle, Left:=610, Top:=4, Width:=160, Height:=27).Name = Split(SlideInfo, "_")(2)
ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.Range(Split(SlideInfo, "_")(2)).Line.Weight = 1.5
ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.Range(Split(SlideInfo, "_")(2)).Line.ForeColor.RGB = RGB(250, 200, 0)
ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.Range(Split(SlideInfo, "_")(2)).Fill.ForeColor.RGB = RGB(255, 255, 255)
ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.Range(Split(SlideInfo, "_")(2)).Fill.BackColor.RGB = RGB(250, 200, 0)
ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.Range(Split(SlideInfo, "_")(2)).TextFrame.TextRange.Text = Split(SlideInfo, "_")(2)
ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.Range(Split(SlideInfo, "_")(2)).TextFrame.TextRange.Font.Color.RGB = RGB(250, 200, 0)
ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.Range(Split(SlideInfo, "_")(2)).TextFrame.TextRange.Font.Size = 14
ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.Range(Split(SlideInfo, "_")(2)).TextFrame.TextRange.Font.Name = "Arial"
ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.Range(Split(SlideInfo, "_")(2)).TextFrame.TextRange.Font.Underline = msoTrue
ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.Range(Split(SlideInfo, "_")(2)).ActionSettings(ppMouseClick).Action = ppActionHyperlink
ActivePresentation.Slides(CInt(Split(SlideInfo, "_")(1))).Shapes.Range(Split(SlideInfo, "_")(2)).ActionSettings(ppMouseClick).Hyperlink.SubAddress = mySlide

Else
    MsgBox ("Â ïðåçåíòàöèè íåò ñëàéäà " + Split(SlideInfo, "_")(1))
End If
End If

End Sub

Sub hi()
    Debug.Print "Start"
End Sub

Sub create_all_link()
On Error Resume Next
f = FreeFile
a = 0
b = 0
If Dir(ActivePresentation.Path & "\title_dict.txt") <> "" Then
    Open ActivePresentation.Path & "\title_dict.txt" For Input As #f
    Do While Not EOF(f)
      Line Input #f, s
      a = a + 1
      If s <> "" Then
        If Split(s, "_")(0) = "" Or Split(s, "_")(1) = "" Or Split(s, "_")(2) = "" Or Split(s, "_")(3) = "" Then
            Debug.Print "Îøèáêà çàïîëíåíèÿ ñòðîêè " & a & " â ôàéëå " & "title_dict.txt"
        Else
            zero = Split(s, "_")(0)
            one = Split(s, "_")(1)
            two = Split(s, "_")(2)
            three = Split(s, "_")(3)
            

               For Each el In Split(one, ",")
                    If CInt(el) <= ActivePresentation.Slides.Count Then
                        For Each shp In ActivePresentation.Slides(CInt(el)).Shapes
                            c = StrComp("link_", Left(shp.Name, 5), 1)
                            If StrComp("link_", Left(shp.Name, 5), 1) = 0 Then
                                ActivePresentation.Slides(CInt(el)).Shapes.Range(shp.Name).Delete
                            End If
                        Next shp
                        
                        ActivePresentation.Slides(CInt(el)).Shapes.AddShape(Type:=msoShapeRectangle, Left:=610, Top:=4, Width:=160, Height:=27).Name = "link_" & el
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).Line.Weight = 1.5
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).Line.ForeColor.RGB = RGB(250, 200, 0)
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).Fill.ForeColor.RGB = RGB(255, 255, 255)
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).Fill.BackColor.RGB = RGB(250, 200, 0)
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).TextFrame.TextRange.Text = zero
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).TextFrame.TextRange.Font.Color.RGB = RGB(250, 200, 0)
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).TextFrame.TextRange.Font.Size = 14
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).TextFrame.TextRange.Font.Name = "Arial"
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).TextFrame.TextRange.Font.Underline = msoTrue
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).ActionSettings(ppMouseClick).Action = ppActionHyperlink
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).ActionSettings(ppMouseClick).Hyperlink.SubAddress = _
                                                                                                 ActivePresentation.Slides(CInt(Split(Split(s, "_")(3), ",")(0))).SlideNumber
                        b = b + 1
                    Else
                        Debug.Print "Â ñòðîêå " & a & " íîìåðà ñëàéäîâ áîëüøå, ÷åì â ïðåçåíòàöèè."
                    End If
               Next el
               
               For Each el In Split(three, ",")
                    If CInt(el) <= ActivePresentation.Slides.Count Then
                        For Each shp In ActivePresentation.Slides(CInt(el)).Shapes
                            c = StrComp("link_", Left(shp.Name, 5), 1)
                            If StrComp("link_", Left(shp.Name, 5), 1) = 0 Then
                                ActivePresentation.Slides(CInt(el)).Shapes.Range(shp.Name).Delete
                            End If
                        Next shp
                        
                        ActivePresentation.Slides(CInt(el)).Shapes.AddShape(Type:=msoShapeRectangle, Left:=610, Top:=4, Width:=160, Height:=27).Name = "link_" & el
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).Line.Weight = 1.5
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).Line.ForeColor.RGB = RGB(250, 200, 0)
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).Fill.ForeColor.RGB = RGB(255, 255, 255)
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).Fill.BackColor.RGB = RGB(250, 200, 0)
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).TextFrame.TextRange.Text = two
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).TextFrame.TextRange.Font.Color.RGB = RGB(250, 200, 0)
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).TextFrame.TextRange.Font.Size = 14
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).TextFrame.TextRange.Font.Name = "Arial"
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).TextFrame.TextRange.Font.Underline = msoTrue
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).ActionSettings(ppMouseClick).Action = ppActionHyperlink
                        ActivePresentation.Slides(CInt(el)).Shapes.Range("link_" & el).ActionSettings(ppMouseClick).Hyperlink.SubAddress = _
                                                                                                 ActivePresentation.Slides(CInt(Split(Split(s, "_")(1), ",")(0))).SlideNumber
                        b = b + 1
                    Else
                        Debug.Print "Â ñòðîêå " & a & " íîìåðà ñëàéäîâ áîëüøå, ÷åì â ïðåçåíòàöèè."
                    End If
               Next el
                
               


            
        End If
      End If
    Loop
    Close f
    MsgBox ("Óñòàíîâëåíî " & b & " ññûëîê")
Else
    MsgBox ("Íå íàéäåí ôàéë ñ íîìåðàìè ñëàéäîâ.")
End If

End Sub

