<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: updatecategorysequence.asp
' AUTHOR: Steve Loar
' CREATED: 09/10/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the display order of a category. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   09/10/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRecreationCategoryId, iSequence

iRecreationCategoryId = CLng(request("categoryid"))
iSequence = CLng(request("sequence"))

sSql = "UPDATE egov_recreation_categories SET sequenceid = " & iSequence & " WHERE recreationcategoryid = " & iRecreationCategoryId
RunSQL sSql

response.write "UPDATED"


%>