<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkinpsecteddate.asp
' AUTHOR: Steve Loar
' CREATED: 07/24/08
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks whether an inspected date is in the future or not. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   07/24/08	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, sResults, iInspectedDate, iDateDiff

If request("inspecteddate") <> "" Then 
	iInspectedDate = CDate(request("inspecteddate"))

	iDateDiff = DateDiff( "d", iInspectedDate, Date() )

	If iDateDiff >= 0 Then
		sResults = "DATE OK"
	Else
		sResults = "FUTURE DATE"
	End If 
Else
	sResults = "FUTURE DATE"
End If 

response.write sResults



%>