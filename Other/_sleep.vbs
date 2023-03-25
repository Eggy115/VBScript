WScript.Echo "wait " & WScript.Arguments.Item(0) & "seconds"
WScript.Echo "Wait START on " & Now()
WScript.Sleep CDbl(WScript.Arguments.Item(0))*1000
WScript.Echo "Wait END on " & Now()
