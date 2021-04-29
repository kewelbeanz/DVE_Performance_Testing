' Duration of time from user submitting MTF catalog search on "Mask" to complete results display
' Starting from Navigation screen - if combining this test with another, include the necessary steps after the creation of WScript Shell object
' to get to navigation screen, whether it is clicking on close button a few times or doing a file exit, etc.
' Create WScript Shell Object to access filesystem.
Set WshShell = WScript.CreateObject("WScript.Shell")

' Setup for logging
dtmStartTime=Timer

' Moving to the CAIM selection 
WshShell.SendKeys "{DOWN 7}"
WshShell.SendKeys "{ENTER}"



' Timer setup
Dim lengthy
lengthy=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output,fileSystemObject, filePath
filePath="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1247_CAIM_tests2\startCAIM.txt"
Set fileSystemObject=CreateObject("Scripting.FileSystemObject")
Set output=fileSystemObject.CreateTextFile(filePath, true)
output.WriteLine("Duration of time from user submitting MTF catalog search on "Mask" to complete results display was "+lengthy+" seconds.")
output.Close


