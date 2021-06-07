<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: managemasterlistupdate.asp
' AUTHOR: Steve Loar
' CREATED: 08/02/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This processes updates to Rye's Renewal Master List.
'
' MODIFICATION HISTORY
' 1.0	08/02/2011	Steve Loar - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSql, iRenewalId, sPermitHolderType, sApplicantFirstName, sApplicantLastName
Dim sApplicantAddress, sApplicantCity, sApplicantState, sApplicantZip
Dim sApplicantPhone, sVehicleLicense

iRenewalId = CLng( request("renewalid") )
sPermitHolderType = "'" & dbsafe( request("permitholdertype") ) & "'"
sApplicantFirstName = "'" & dbsafe( request("applicantfirstname") ) & "'"
sApplicantLastName = "'" & dbsafe( request("applicantlastname") ) & "'"
sApplicantAddress = "'" & dbsafe( request("applicantaddress") ) & "'"
sApplicantCity = "'" & dbsafe( request("applicantcity") ) & "'"
sApplicantState = "'" & dbsafe( UCase(request("applicantstate")) ) & "'"
sApplicantZip = "'" & dbsafe( request("applicantzip") ) & "'"

If request("applicantphone") <> "" Then 
	sApplicantPhone = "'" & dbsafe( request("applicantphone") ) & "'"
Else
	sApplicantPhone = "NULL"
End If 

If request("vehiclelicense") <> "" Then 
	sVehicleLicense = "'" & dbsafe( UCase(request("vehiclelicense")) ) & "'"
Else
	sVehicleLicense = "NULL"
End If 

sSql = "UPDATE egov_ryepermitrenewals SET "
sSql = sSql & "permitholdertype = " & sPermitHolderType & ", "
sSql = sSql & "applicantfirstname = " & sApplicantFirstName & ", "
sSql = sSql & "applicantlastname = " & sApplicantLastName & ", "
sSql = sSql & "applicantaddress = " & sApplicantAddress & ", "
sSql = sSql & "applicantcity = " & sApplicantCity & ", "
sSql = sSql & "applicantstate = " & sApplicantState & ", "
sSql = sSql & "applicantzip = " & sApplicantZip & ", "
sSql = sSql & "applicantphone = " & sApplicantPhone & ", "
sSql = sSql & "vehiclelicense = " & sVehicleLicense & " "
sSql = sSql & "WHERE renewalid = " & iRenewalId

'response.write sSql & "<br /><br />"

RunSQLStatement sSql

'take them to the master list page
response.redirect "managemasterlist.asp?s=u&applicantfirstnamesearch=" & request("applicantfirstname") & "&applicantlastnamesearch=" & request("applicantlastname") & "&permitholdertypesearch=" & request("permitholdertype")

%>
