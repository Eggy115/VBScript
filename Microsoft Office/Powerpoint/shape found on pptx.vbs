
' List the Shapes Found on a Microsoft PowerPoint Slide



Const ppLayoutText = 2

Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True
Set objPresentation = objPPT.Presentations.Add
objPresentation.ApplyTemplate _
    ("C:\Program Files\Microsoft Office\" & _
        "Templates\Presentation Designs\Globe.pot")
Set objSlide = objPresentation.Slides.Add(1, ppLayoutText)

Set objShapes = objSlide.Shapes

For Each objShape in objShapes
    Wscript.Echo objShape.Name
Next
