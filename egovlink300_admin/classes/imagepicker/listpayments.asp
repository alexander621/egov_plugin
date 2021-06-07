<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<link href="../global.css" rel="stylesheet" type="text/css">
</HEAD>

<BODY>
<p><% Call subListForms() %></p>
</BODY>
</HTML>


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' SUB SUBLISTFORMS()
'------------------------------------------------------------------------------------------------------------
Sub subListForms()

	sSQL = "SELECT DISTINCT paymentserviceid,paymentservicename FROM egov_paymentservices LEFT OUTER JOIN egov_organizations_to_paymentservices ON egov_paymentservices.paymentserviceid=egov_organizations_to_paymentservices.paymentservice_id where (egov_organizations_to_paymentservices.paymentservice_enabled <> 0 and (egov_organizations_to_paymentservices.orgid=" & session("orgid") & " OR egov_organizations_to_paymentservices.orgid=0))"
	Set oFormList = Server.CreateObject("ADODB.Recordset")
	oFormList.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oFormList.EOF Then

		response.write "<table cellspacing=0 cellpadding=2 >"

		Do while NOT oFormList.EOF 
			
			sFormNameOnly = replace(UCASE(oFormList("paymentservicename")),"'","\'")
			sFormNameOnly = replace(sFormNameOnly,chr(10),"")
			sFormNameOnly = replace(sFormNameOnly,chr(13),"")
			sFormNameOnly = Trim(sFormNameOnly)

			response.write "<tr style=""cursor:hand;"" onClick=""parent.document.frmPaymentLink.AFormName.value='" & sFormNameOnly  & "';parent.document.frmPaymentLink.iFormID.value=" & oFormList("paymentserviceid") & """ ><td > (" & oFormList("paymentserviceid") & ") </td><td>" & UCASE(oFormList("paymentservicename")) & "</td></tr>" 
			oFormList.MoveNext
		Loop

		response.write "</table>"

	End If
	Set oFormList = Nothing 

End Sub
%>