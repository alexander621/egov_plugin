<!-- #include file="../includes/common.asp" //-->
<%
'Determine if this is an Ajax Routine
 lcl_isAjaxRoutine = "N"

 if request("isAjaxRoutine") = "Y" then
    lcl_isAjaxRoutine = UCASE(request("isAjaxRoutine"))
 end if

 Call subChangeDisplay(request("iMembershipID"),request("sResidentType"), request("public_display"), request("sMembershipType"), lcl_isAjaxRoutine)

'------------------------------------------------------------------------------
' SUB subChangeDisplay( sResidentType, public_display )
' AUTHOR: Steve Loar
' CREATED: 02/05/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' 02/05/06  Steve Loar - Initial Version
' 02/25/09  David Boyer - Added "isAjaxRoutine" parameter and pass-backs.
'------------------------------------------------------------------------------
sub subChangeDisplay(iMembershipId, sResidentType, public_display, p_membershiptype, iIsAjaxRoutine)

 'First delete the existing record
  sSQL = "DELETE FROM egov_membership_rate_displays "
  sSQL = sSQL & " WHERE membershipid = " & iMembershipId
  sSQL = sSQL & " AND resident_type = '" & sResidentType & "'"

  set oDeleteDisplay = Server.CreateObject("ADODB.Recordset")
  oDeleteDisplay.Open sSQL, Application("DSN"), 3, 1

  set oDeleteDisplay = nothing

 'Insert a new record
  sSQL = "INSERT INTO egov_membership_rate_displays (membershipid, resident_type, public_display) VALUES ( "
  sSQL = sSQL & iMembershipId & ", "
  sSQL = sSQL & "'" & sResidentType & "', "
  sSQL = sSQL & public_display
  sSQL = sSQL & ")"

  set oUpdateDisplay = Server.CreateObject("ADODB.Recordset")
  oUpdateDisplay.Open sSQL, Application("DSN"), 3, 1

  set oUpdateDisplay = nothing
 	'Set oCmd = Server.CreateObject("ADODB.Command")
 	'With oCmd
	 '	.ActiveConnection = Application("DSN")
	 	'Delete any old record
 	'	.CommandText = "DELETE FROM egov_membership_rate_displays WHERE membershipid = " & iMembershipId & " AND resident_type = '" & sResidentType & "'"
 	'	.Execute
 	'	.CommandText = sSql
 	'	.Execute
 	'End With
 	'Set oCmd = Nothing

'Determine if this is an Ajax call or redirect
 if iIsAjaxRoutine = "Y" then
    response.write "Successfully Updated..."
 else
   	response.redirect( "poolpass_rates.asp?sResidentType=" & sResidentType & "&sMembershipType=" & p_membershiptype)
 end if

end sub
%>