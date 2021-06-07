<!-- #include file="includes/JSON_2.0.2.asp" //-->
<!-- #include file="includes/common.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitrenewallookup.asp
' AUTHOR: Steve Loar
' CREATED: 07/21/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Gets the renewal information for a Rye Permit Renewal, or waitlist.
'				Called via AJAX from payment.asp
'
' MODIFICATION HISTORY
' 1.0   07/21/2011	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sResponse, sSql, oRs, sPermitHolderType, sApplicantFirstName, sApplicantLastName

' Create the JSON object to pass data back to the calling page
Set sResponse = jsObject()

sPermitHolderType = request("permitholdertype")
sApplicantFirstName = request("applicantfirstname")
sApplicantLastName = request("applicantlastname")


sSql = "SELECT renewalid, ISNULL(applicantfirstname, '') AS applicantfirstname, ISNULL(applicantlastname,'') AS applicantlastname, "
sSql = sSql & "ISNULL(applicantaddress, '') AS applicantaddress, ISNULL(applicantcity,'') AS applicantcity, "
sSql = sSql & "ISNULL(applicantstate,'') AS applicantstate, ISNULL(applicantzip,'') AS applicantzip, "
sSql = sSql & "ISNULL(applicantphone,'') AS applicantphone, ISNULL(vehiclelicense,'') AS vehiclelicense, ISNULL(vehicle2license,'') as vehicle2license, hasrenewed "
sSql = sSql & "FROM egov_ryepermitrenewals "
sSql = sSql & "WHERE permitholdertype = '" & Track_DBsafe( sPermitHolderType ) & "' "
sSql = sSql & "AND LOWER(applicantfirstname) = '" & LCase(Track_DBsafe( sApplicantFirstName )) & "' "
sSql = sSql & "AND LOWER(applicantlastname) = '" & LCase(Track_DBsafe( sApplicantLastName )) & "' "
sSql = sSql & "AND year = " & Year(Now()) & " "
sSql = sSql & "AND orgid = " & iOrgID
sSql = sSql & " ORDER BY hasrenewed,renewalid"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then 
	If oRs("hasrenewed") Then
		sResponse("flag") = "duplicate"
	Else
		sResponse("flag") = "success"
		sResponse("renewalid") = Trim(oRs("renewalid"))
		sResponse("applicantfirstname") = Trim(oRs("applicantfirstname"))
		sResponse("applicantlastname") = Trim(oRs("applicantlastname"))
		sResponse("applicantaddress") = Trim(oRs("applicantaddress"))
		sResponse("applicantcity") = Trim(oRs("applicantcity"))
		sResponse("applicantstate") = Trim(oRs("applicantstate"))
		sResponse("applicantzip") = Trim(oRs("applicantzip"))
		sResponse("applicantphone") = Trim(oRs("applicantphone"))
		sResponse("vehicle1license") = Trim(oRs("vehiclelicense"))
		sResponse("vehicle2license") = Trim(oRs("vehicle2license"))
	End If 
Else
	sResponse("flag") = "notfound"
	'sResponse("sql") = sSql
End If 


sResponse.Flush
%>
