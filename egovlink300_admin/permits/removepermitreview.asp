<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: removepermitreview.asp
' AUTHOR: Steve Loar
' CREATED: 06/23/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Removes a review from a permit
'
' MODIFICATION HISTORY
' 1.0   06/23/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitReviewId, sSql

iPermitReviewId = CLng(request("permitreviewid"))

' Remove from the permit review table
sSql = "DELETE FROM egov_permitreviews WHERE permitreviewid = " & iPermitReviewId
RunSQL sSql

response.write "Success"

%>
