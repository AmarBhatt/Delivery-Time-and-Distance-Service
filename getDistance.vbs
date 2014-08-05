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

if (UBound(arrFields) = 1) then

	origin = Replace(arrFields(0)," ","+")
	destination = Replace(arrFields(1)," ","+")


	' Call Google Maps API --> https://developers.google.com/maps/documentation/distancematrix/
	Dim o
	Set o = CreateObject("MSXML2.XMLHTTP")
	o.open "GET", "http://maps.googleapis.com/maps/api/distancematrix/xml?origins=" & origin & "&destinations=" & destination & "&mode=driving&language=en-EN", False
	o.send
	' o.responseText now holds the response as a string.

	' Process the XML output to get duration and distance
	Dim xmlDoc
	dim durItem, duration, distItem, distance
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	xmlDoc.async = False 
	xmlDoc.LoadXml(o.responseText)

	' Check if request returned valid
	Set statusList = xmlDoc.getElementsByTagName("element")

	For each statItem in statusList
		Set status = statItem.SelectSingleNode("status")
	Next


	if (StrComp(status.text,"OK") = 0) then
		' Get the driving time between the two addresses
		Set durationList = xmlDoc.getElementsByTagName("duration")
		
		For each durItem in durationList
			Set duration = durItem.SelectSingleNode("text")
		Next
		
		' Get the driving distance between the two addresses
		Set distanceList = xmlDoc.getElementsByTagName("distance")
		
		For each distItem in distanceList
			Set distance = distItem.SelectSingleNode("value")
		Next
		
		' Write results to disk
		Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(fso.GetAbsolutePathName(".") & "\addresses.txt",8,false)
		objFileToWrite.Write("," & duration.text & "," & Round(distance.text/1609.344, 2) & " miles")
		objFileToWrite.Close
		Set objFileToWrite = Nothing
	else
		WScript.Echo "An Error Occured: ERROR_STATUS = " & status.text
	end if
else
	WScript.Echo "Origin and Destination addresses required. Comma deliminated."
end if