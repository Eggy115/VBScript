dim sURL
sURL= "https://www.github.com/Eggy115"
dim objIE
Set objIE = CreateObject("InternetExplorer.Application")   
objIE.visible = True 
objIE.navigate(surl)
