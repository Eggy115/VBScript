' Modify System Restore Configuration Values


Const GLOBAL_INTERVAL_IN_SECONDS = 100000
Const LIFE_INTERVAL_IN_SECONDS = 8000000
Const SESSION_INTERVAL_IN_SECONDS = 500000
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")

Set objItem = objWMIService.Get("SystemRestoreConfig='SR'")
objItem.DiskPercent = 10
objItem.RPGlobalInterval = GLOBAL_INTERVAL_IN_SECONDS
objItem.RPLifeInterval = LIFE_INTERVAL_IN_SECONDS
objItem.RPSessionInterval = SESSION_INTERVAL_IN_SECONDS
objItem.Put_

