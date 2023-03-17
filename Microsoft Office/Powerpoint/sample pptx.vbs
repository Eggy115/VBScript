' Create a Sample Microsoft PowerPoint Presentation



Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True
Set objPresentation = objPPT.Presentations.Add
objPresentation.ApplyTemplate("C:\Program Files\Microsoft Office\Templates\Presentation Designs\Globe.pot")

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")

For Each objProcess in colProcesses
    Set objSlide = objPresentation.Slides.Add(1,2)
    Set objShapes = objSlide.Shapes

    Set objTitle = objShapes.Item("Rectangle 2")
    objTitle.TextFrame.TextRange.Text = objProcess.Name

    strText = "Working set size: " & objProcess.WorkingSetSize & vbCrLf
    strText = strText & "Priority: " & objProcess.Priority & vbCrLf
    strText = strText & "Thread count: " & objProcess.ThreadCount & vbCrLf

    Set objTitle = objShapes.Item("Rectangle 3")
    objTitle.TextFrame.TextRange.Text = strText
Next

objPresentation.SaveAs("C:\Scripts\Process.ppt")
objPresentation.Close
objPPT.Quit
