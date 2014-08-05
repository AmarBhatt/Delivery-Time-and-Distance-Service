' Get the file path
dim fso: set fso = CreateObject("Scripting.FileSystemObject")

' Read comma deliminated file with origin and destination addresses
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(fso.GetAbsolutePathName(".") & "\addresses.txt",1)
Do Until objFileToRead.AtEndOfStream
    strLine = objFileToRead.ReadLine
    arrFields = Split(strLine, ",")
Loop

objFileToRead.Close
Set objFileToRead = Nothing

' Check if duration and distance enteries exist
if(UBound(arrFields) = 3) then
	' Put reading code here
	WScript.Echo arrFields(2)
	WScript.Echo arrFields(3)
else
	WScript.Echo "Driving time and distance not found for this entry"
end if
