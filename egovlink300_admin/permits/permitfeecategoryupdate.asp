<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitfeecategoryupdate.asp
' AUTHOR: Steve Loar
' CREATED: 12/17/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates the permit fee categories surcharges
'
' MODIFICATION HISTORY
' 1.0   12/17/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, x, iApply, sLabel, sRate

For x = CLng(request("min")) To CLng(request("max"))
	If request("permitfeecategorytypeid" & x) <> "" Then 
		If request("apply" & x) = "on" Then
			iApply = 1
		Else
			iApply = 0
		End If 
		If request("label" & x) = "" Then
			sLabel = "NULL" 
		Else
			sLabel = "'" & dbsafe(request("label" & x)) & "'"
		End If 
		If request("rate" & x) = "" Then 
			sRate = "NULL"
		Else
			sRate = request("rate" & x)
		End If 
		sSql = "UPDATE egov_permitfeecategorytypes SET surchargelabel = " & sLabel & ", surchargerate = " & sRate & ", applysurcharge = " & iApply & " WHERE orgid = " & session("orgid") & " AND permitfeecategorytypeid = " & x
		RunSQL sSql 
	End If 
Next 

response.redirect "permitfeecategories.asp"


'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

%>