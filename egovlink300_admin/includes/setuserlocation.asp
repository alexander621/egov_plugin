<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: setuserlocation.asp
' AUTHOR: Steve Loar	
' CREATED: 02/20/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the user location for this session
'
' MODIFICATION HISTORY
' 1.0   02/20/2007   Steve Loar - Initial code 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

' Set the session variable to the new location
Session("LocationId") = request("locationid")

' Update the cookie to have the new value
Response.Cookies("User")("LocationId") = Session("LocationId")

' Return 
response.write Session("LocationId")

%>