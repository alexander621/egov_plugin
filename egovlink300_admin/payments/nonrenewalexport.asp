<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: nonrenewalexport.asp
' AUTHOR: Steve Loar
' CREATED: 08/02/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Export of the Rye master list showing those who did not renew
'
' MODIFICATION HISTORY
' 1.0   08/02/2011	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, sDate

' SET UP PAGE OPTIONS
sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=NonrewalMasterList_" & sDate & ".xls"

sSql = "SELECT permitholdertype, applicantfirstname, applicantlastname, ISNULL(applicantaddress,'') AS applicantaddress, "
sSql = sSql & "ISNULL(applicantcity,'') AS applicantcity, ISNULL(applicantstate,'') AS applicantstate, "
sSql = sSql & "ISNULL(applicantzip,'') AS applicantzip, ISNULL(applicantphone,'') AS applicantphone, "
sSql = sSql & "ISNULL(vehiclelicense,'') AS vehiclelicense, ISNULL(rr,'') AS rr, ISNULL(hc,'') AS hc, ISNULL(emailaddress,'') as emailaddress "
sSql = sSql & "FROM egov_ryepermitrenewals "
sSql = sSql & "WHERE hasrenewed = 0 "
sSql = sSql & "AND orgid = " & session("orgid") & " AND year = " & Year(Now()) 
sSql = sSql & " ORDER BY renewalid"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

If Not oRs.EOF Then
	response.write vbcrlf & "<html><body>"
	
	'response.write sSql & "<br /><br />"

	response.write vbcrlf & "<table border=""1"" cellpadding=""4"">"

	response.write vbcrlf & "<tr>"
	response.write "<td align=""center""><b>Permit Holder Type</b></td>"
	response.write "<td align=""center""><b>Applicant First Name</b></td>"
	response.write "<td align=""center""><b>Applicant Last Name</b></td>"
	response.write "<td align=""center""><b>Applicant Address</b></td>"
	response.write "<td align=""center""><b>Applicant City</b></td>"
	response.write "<td align=""center""><b>Applicant State</b></td>"
	response.write "<td align=""center""><b>Applicant Zip</b></td>"
	response.write "<td align=""center""><b>Applicant Phone</b></td>"
	response.write "<td align=""center""><b>Vehicle License</b></td>"
	response.write "<td align=""center""><b>RR</b></td>"
	response.write "<td align=""center""><b>H/C</b></td>"
	response.write "<td align=""center""><b>Applicant Email</b></td>"
	response.write "</tr>"
	response.flush
		
	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"
		response.write "<td>" & oRs("permitholdertype") & "</td>"
		response.write "<td>" & oRs("applicantfirstname") & "</td>"
		response.write "<td>" & oRs("applicantlastname") & "</td>"
		response.write "<td>" & oRs("applicantaddress") & "</td>"
		response.write "<td>" & oRs("applicantcity") & "</td>"
		response.write "<td>" & oRs("applicantstate") & "</td>"
		response.write "<td>" & oRs("applicantzip") & "</td>"
		response.write "<td>" & oRs("applicantphone") & "</td>"
		response.write "<td>" & oRs("vehiclelicense") & "</td>"
		response.write "<td>" & oRs("rr") & "</td>"
		response.write "<td>" & oRs("hc") & "</td>"
		response.write "<td>" & oRs("emailaddress") & "</td>"
		response.write "</tr>"
		response.flush

		oRs.MoveNext
	Loop

	response.write vbcrlf & "</table></body></html>"
	response.flush

End If 
		
oRs.Close
Set oRs = Nothing 

%>
