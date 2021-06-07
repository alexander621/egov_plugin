<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitcustomfieldtypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 10/22/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit custom field types
'
' MODIFICATION HISTORY
' 1.0   10/22/2010   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iCustomFieldTypeid, sSql, iFieldTypeId, sFieldName, sPdfFieldName, sPrompt, sFieldSize
Dim sValueList, sReportTitle

iCustomFieldTypeid = CLng(request("cft"))

iFieldTypeId = request("fieldtypeid")

sFieldName = "'" & dbsafe(request("fieldname")) & "'"

sReportTitle = "'" & dbsafe(request("reporttitle")) & "'"

If request("pdffieldname") <> "" Then
	sPdfFieldName = "'" & dbsafe(Replace(LCase(request("pdffieldname"))," ","")) & "'"
Else
	sPdfFieldName = "NULL"
End If 

sPrompt = "'" & dbsafe(request("prompt")) & "'"

If request("valuelist") <> "" Then
	sValueList = "'" & dbsafe(request("valuelist")) & "'"
Else
	sValueList = "NULL"
End If 
'response.write "sValueList = " & sValueList & "<br /><br />"

If request("fieldsize") <> "" Then
	If clng(request("fieldsize")) = clng(0) Then
		sFieldSize = "NULL"
	Else 
		sFieldSize = request("fieldsize")
	End If 
Else
	sFieldSize = "NULL"
End If 

If iCustomFieldTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_permitcustomfieldtypes ( orgid, fieldtypeid, fieldname, pdffieldname, prompt, valuelist, "
	sSql = sSql & "fieldsize, reporttitle, isactive ) VALUES ( "
	sSql = sSql & session("orgid") & ", " & iFieldTypeId & ", " & sFieldName & ", " & sPdfFieldName & ", " & sPrompt & ", "
	sSql = sSql & sValueList & ", " & sFieldSize & ", " & sReportTitle & ", 1 )"
	'response.write sSql & "<br /><br />"

	iCustomFieldTypeid = RunIdentityInsert( sSql )
	
	sSuccessMsg = "This Custom Field Type has been created."
Else 
	sSql = "UPDATE egov_permitcustomfieldtypes SET "
	sSql = sSql & "fieldtypeid = " & iFieldTypeId & ", "
	sSql = sSql & "fieldname = " & sFieldName & ", "
	sSql = sSql & "pdffieldname = " & sPdfFieldName & ", "
	sSql = sSql & "prompt = " & sPrompt & ", "
	sSql = sSql & "valuelist = " & sValueList & ", "
	sSql = sSql & "fieldsize = " & sFieldSize & ", "
	sSql = sSql & "reporttitle = " & sReportTitle & " "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND customfieldtypeid = " & iCustomFieldTypeid
	'response.write sSql & "<br /><br />"

	RunSQL sSql 

	sSuccessMsg = "Your changes have been saved."
End If 

'response.write sSuccessMsg & "<br /><br />"

response.redirect "permitcustomfieldtypeedit.asp?cft=" & iCustomFieldTypeid & "&success=" & sSuccessMsg

%>
