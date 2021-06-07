<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: header_update.asp
' AUTHOR: Steve Loar
' CREATED: 05/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates or updates the Receipt Headers
'
' MODIFICATION HISTORY
' 1.0   05/15/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iInvoiceHeaderDisplayId, sInvoiceHeader, sSql, sLogoURL, iLogoURLDisplayId

iInvoiceHeaderDisplayId = request("headerdisplayid")
sInvoiceHeader = DBsafeWithHTML(request("invoiceheader"))

iLogoURLDisplayId = request("logodisplayid")
sLogoURL = DBsafeWithHTML(request("logourl"))

' Clean out the old header 
sSql = "Delete from egov_organizations_to_displays where orgid = " & Session("orgid") & " and displayid = " & iInvoiceHeaderDisplayId
RunSQL sSql 

' Clean out the old logo 
sSql = "Delete from egov_organizations_to_displays where orgid = " & Session("orgid") & " and displayid = " & iLogoURLDisplayId
RunSQL sSql 



' New Header
sSql = "Insert into egov_organizations_to_displays ( orgid, displayid, displaydescription ) values ( "
sSql = sSql & Session("orgid") & ", " & iInvoiceHeaderDisplayId & ", '" & sInvoiceHeader & "' )"
RunSQL sSql 

' New Header
sSql = "Insert into egov_organizations_to_displays ( orgid, displayid, displaydescription ) values ( "
sSql = sSql & Session("orgid") & ", " & iLogoURLDisplayId & ", '" & sLogoURL & "' )"
RunSQL sSql 


response.redirect "header_edit.asp"


'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

%>