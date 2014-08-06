' Testing --> http://maps.googleapis.com/maps/api/distancematrix/xml?origins=125+Ormsby+Dr+Syracuse+NY+13219&destinations=60+Lomb+Memorial+Drive+NY+13219&mode=driving&language=en-EN


' Get the file path
dim fso: set fso = CreateObject("Scripting.FileSystemObject")

' Read comma deliminated file with origin and destination addresses
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(fso.GetAbsolutePathName(".") & "\addresses.csv",1)
Do Until objFileToRead.AtEndOfStream
    strLine = objFileToRead.ReadLine
    arrFields = Split(strLine, ",")
Loop

objFileToRead.Close
Set objFileToRead = Nothing

if (UBound(arrFields) = 1) then
	
	origin = Replace(arrFields(0),"""","")
	destination = Replace(arrFields(1),"""","")
	originMod = Replace(origin," ","+")
	destinationMod = Replace(destination," ","+")

	' Call Google Maps API --> https://developers.google.com/maps/documentation/distancematrix/
	Dim o
	Set o = CreateObject("MSXML2.XMLHTTP")
	o.open "GET", "http://maps.googleapis.com/maps/api/distancematrix/xml?origins=" & originMod & "&destinations=" & destinationMod & "&mode=driving&language=en-EN", False
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

		Dim orig, dest
		'Get origin address used by Google
		Set originList = xmlDoc.getElementsByTagName("DistanceMatrixResponse")		
		For each origItem in originList
			Set orig = origItem.SelectSingleNode("origin_address")
		Next
		originAddress = Replace(orig.text,",","")
		originAddress = Replace(originAddress," USA","")

		'Get destination address used by Google
		Set destinationList = xmlDoc.getElementsByTagName("DistanceMatrixResponse")		
		For each destItem in destinationList
			Set dest = destItem.SelectSingleNode("destination_address")
		Next
		destinationAddress = Replace(dest.text,",","")
		destinationAddress = Replace(destinationAddress," USA","")

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

		' Figure out how the distance was calculated by checking the first 3 characters of each address
		if((StrComp(Mid(originAddress,1,3),Mid(origin,1,3)) = 0) And (StrComp(Mid(destinationAddress,1,3),Mid(destination,1,3)) = 0)) then
			fromWhere = "address"
		else
			fromWhere = "zipcode"
		end if

		' Write results to disk
		Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(fso.GetAbsolutePathName(".") & "\distanceResult.csv",2,true)
		objFileToWrite.Write("""" & originAddress & """" & "," & """" & destinationAddress & """" & "," & """" & Round(distance.text/1609.344, 2) & """" & "," & """" & fromWhere & """")
		objFileToWrite.Close
		Set objFileToWrite = Nothing
	else
		error = status.text
		' Write error results to disk
		Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(fso.GetAbsolutePathName(".") & "\distanceResult.csv",2,true)
		objFileToWrite.Write("""" & origin & """" & "," & """" & destination & """" & "," & """" & "SERVICE RETURNED AN ERROR" & """" & "," & """" & error & """")
		objFileToWrite.Close
		Set objFileToWrite = Nothing
	end if
else
	' Write error results to disk
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(fso.GetAbsolutePathName(".") & "\distanceResult.csv",2,true)
	objFileToWrite.Write("""" & "Origin and Destination addresses required. Comma deliminated."&"""")
	objFileToWrite.Close
	Set objFileToWrite = Nothing
end if