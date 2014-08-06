Delivery-Time-and-Distance-Service
==================================

A time and distance estimator between two addresses using VBScript and Google Maps Distance Matrix API


I was asked to solve the problem of getting an accurate distance for two addresses for delivery planning.  This needed to be done without relying on external dependencies or libraries.  It also needed to be run from the command line, which needed to read and write a file.  I chose VBScript because I was solving for a someone whose clients all ran on windows machines. Also, VBScript works well alongside the programs he has created.

This uses the google maps distance matrix api (https://developers.google.com/maps/documentation/distancematrix/). 

This entire service is not dependent on any external service or library. It can be run on any Windows command line.  

The getDistance code reads in a file of two comma deliminated addresses (addresses.csv).  These are written in the order of origin and destination.  Once these addresses are processed by google maps using an http request, the result is written to the distanceResult.csv file. The format for the output is shown below:

* if success: "ORIGIN USED BY GOOGLE","DESTINATION USED BY GOOGLE","MILES","ADDRESS/ZIPCODE"
  * I am checking whether or not Google is reading the address or doing it by zipcode (Google does this for us).  Basically, if the address is bogus it will use the zipcode and you will see in the outputted addresses an address that just has city state and zip instead of the address.  However, if the zipcode is bogus but the address is sound it will fix the zipcode/state for you. Hence the last field "ADDRESS/ZIPCODE" represents how it was proccessed by Google.

* if bad response: "INPUT ORIGIN","INPUT DESTINATION","ERROR","ERROR MESSAGE FROM GOOGLE"

* if bad input: "ERROR MESSAGE"

I have commented the code, as well as put error checks in it. Also, the outputs for the code give you the estimated driving time as well as the estimated driving distance in miles (driving time is implemented in code but not shown in output). 

Let me know if you have any questions or requests! I'd be happy to help :)

