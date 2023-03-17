' Run a Microsoft PowerPoint Slide Show



Const ppAdvanceOnTime = 2
Const ppShowTypeKiosk = 3
Const ppSlideShowDone = 5

Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True

Set objPresentation = objPPT.Presentations.Open("C:\Scripts\Process.ppt")

objPresentation.Slides.Range.SlideShowTransition.AdvanceTime = 2
objPresentation.Slides.Range.SlideShowTransition.AdvanceOnTime = TRUE

objPresentation.SlideShowSettings.AdvanceMode = ppAdvanceOnTime 
objPresentation.SlideShowSettings.ShowType = ppShowTypeKiosk
objPresentation.SlideShowSettings.StartingSlide = 1
objPresentation.SlideShowSettings.EndingSlide = _
    objPresentation.Slides.Count

Set objSlideShow = objPresentation.SlideShowSettings.Run.View

Do Until objSlideShow.State = ppSlideShowDone
Loop
