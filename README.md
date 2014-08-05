Delivery-Time-and-Distance-Service
==================================

A time and distance estimator between two addresses using VBScript and Google Maps Distance Matrix API


I was asked to solve the problem of getting an accurate distance for two addresses for delivery planning.  This needed to be done without relying on external dependencies or libraries.  It also needed to be run from the command line, which needed to read and write a file.  I chose VBScript because I was solving for a customer whose clients all ran windows machines.  My customer also writes code in VBS and C, so he will be able to maintain and add to the code I have written.

This uses the google maps distance matrix api (https://developers.google.com/maps/documentation/distancematrix/). 

This entire service is not dependent on any external service or library. It can be run on any Windows command line.  

The getDistance code reads in a file of two comma deliminated addresses (addresses.txt).  These are written in the order of origin and destination.  Once these addresses are processed by google maps using an http request, the result is concatenated to the addresses.txt file.  You can then call the readDistanceResult code.  That code is left as a layout for you to modify to do what you want with the outputs.  Currently it just outputs to the screen a message prompt with the values.

I have commented the code, as well as put error checks in it. Also, the outputs for the code give you the estimated driving time as well as the estimated driving distance in miles. 

Let me know if you have any questions or requests! I'd be happy to help :)

