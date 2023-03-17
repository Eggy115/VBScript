' Modify a Bulleted List in Microsoft PowerPoint



Const ppLayoutText = 2

Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True
Set objPresentation = objPPT.Presentations.Add
objPresentation.ApplyTemplate _
    ("C:\Program Files\Microsoft Office\" & _
        "Templates\Presentation Designs\Globe.pot")
Set objSlide = objPresentation.Slides.Add _
    (1, ppLayoutText)

Set objShapes = objSlide.Shapes

strText = "a" & vbCrLf
strText = strText & "b" & vbcrlf
strText = strtext & "c"

Set objTitle = objShapes.Item("Rectangle 3")
objTitle.TextFrame.TextRange.Text = strText
