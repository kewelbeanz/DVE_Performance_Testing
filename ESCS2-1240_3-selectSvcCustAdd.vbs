' Modify privileges select svc customer add

' Create WScript object to access filesystem
Set WshShell=WScript.CreateObject("WScript.Shell")

' Start Timer object to count milliseconds
dmlssStartTime=Timer

' To navigate to assignment other than assemblage management the default
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"

' Going to Vertical toolbar
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"  ' Add role cust support window should display

WshShell.SendKeys "SAFETY REP WORK REQUEST"  ' Provided the role name box is active
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{DOWN 16}" ' Should be in read box for FM WORK REQUEST
WshShell.SendKeys " "  ' Should select first checkbox 
WshShell.SendKeys "{DOWN}"  ' Focus should be on Save button - try TAB also
WshShell.SendKeys "{ENTER}"

' lengthy contains time as determined by Timer
Dim lengthy
lengthy=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output,fileSystemObject, filePath
filePath="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1240_System_Services\selectSvcCustAdd.txt"
Set fileSystemObject=CreateObject("Scripting.FileSystemObject")
Set output=fileSystemObject.CreateTextFile(filePath, true)
output.WriteLine("Modify privileges select svc customer add took "+lengthy+" seconds.")
output.Close
