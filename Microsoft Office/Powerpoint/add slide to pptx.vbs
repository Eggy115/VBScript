
' Add a Slide to a Microsoft PowerPoint Presentation



Const ppLayoutText = 2

Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True

Set objPresentation = objPPT.Presentations.Add

objPresentation.ApplyTemplate _
    ("C:\Program Files\Microsoft Office\" & _
        "Templates\Presentation Designs\Globe.pot")

Set objSlide = objPresentation.Slides.Add _
    (1, ppLayoutText)
