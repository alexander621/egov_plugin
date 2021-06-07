<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitcustomfieldtypedelete.asp
' AUTHOR: Steve Loar
' CREATED: 10/22/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the permit custom field types
'
' MODIFICATION HISTORY
' 1.0   10/22/2010   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iCustomFieldTypeid, sSql

iCustomFieldTypeid = CLng(request("cft"))

' flag it as deleted
sSql = "UPDATE egov_permitcustomfieldtypes SET isactive = 0 "
sSql = sSql & "WHERE orgid = " & session("orgid") & " AND customfieldtypeid = " & iCustomFieldTypeid

RunSQL sSql


' Remove it from the permit types that have it.
sSql = "DELETE FROM egov_permittypes_to_permitcustomfieldtypes WHERE customfieldtypeid = " & iCustomFieldTypeid

RunSQL sSql

response.redirect "permitcustomfieldtypelist.asp?success=cftr"

%>

