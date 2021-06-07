<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getdaysinmonth.asp
' AUTHOR: Steve Loar
' CREATED: 08/20/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Gets the days in the passed month for the year 2009. Called via AJAX.
'
' MODIFICATION HISTORY
' 1.0   08/20/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMonth, sDayPickName, x, dStartDate, iEnd, dEndDate, sResults

iMonth = clng(request("imonth"))
sDayPickName = request("spickname")

' We will try using the 2009 dates since this is not a leap year, or any other funny thing
dStartDate = CDate( iMonth & "/1/2009")
dEndDate = DateAdd("m", 1, dStartDate)
dEndDate = DateAdd("d", -1, dEndDate)
iEnd = Day(dEndDate)

sResults = "<select id=""" & sDayPickName & """ name=""" & sDayPickName & """>"
For x = 1 To iEnd
	sResults = sResults & vbcrlf & "<option value=""" & x & """>" & x & "</option>"
Next
sResults = sResults & vbcrlf & "</select>"

response.write sResults

%>
