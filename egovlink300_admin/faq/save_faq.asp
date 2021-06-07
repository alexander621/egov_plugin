<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../action_line/action_line_global_functions.asp" //-->
<%
Call subUpdateFaqs(request("ifaqid"), session("orgid"), session("userid"), request("faqtype"), request("FAQCategoryId"), request("FaqQ"), _
                   request("FaqA"), request("publicationstart"),request("publicationend"), request("sendTo_RSS"), request("requestid"))

'------------------------------------------------------------------------------
sub subUpdateFaqs(iFaqID, iOrgID, iUserID, iFAQType, iFAQCategoryId, iFaqQ, iFaqA, iPublicationStart, iPublicationEnd, _
                  iSendToRSS, iRequestID)
	Dim oCmd, sFAQCategoryId, sWhere, sPubStartDate, sPubEndDate, iIsAjaxRoutine
	
	if CLng(iFAQCategoryId) = CLng(0) then
  		sFAQCategoryId = "NULL"
 else
		  sFAQCategoryId = CLng(iFAQCategoryID)
 end if

 if iPublicationStart = "" then
  		sPubStartDate = "NULL"
 else
  		sPubStartDate = "'" & iPublicationStart & "'"
 end if

 if iPublicationEnd = "" then
  		sPubEndDate = "NULL"
 else
		  sPubEndDate = "'" & iPublicationEnd & "'"
 end if

 if iFAQType = "" then
    iFAQType = "FAQ"
 end if

'-- UPDATE --------------------------------------------------------------------
 if iFaqID <> "" then
'------------------------------------------------------------------------------
  		sSQL = "UPDATE FAQ SET "
    sSQL = sSQL & "FaqQ = '"            & DBsafe(iFaqQ)   & "', "
    sSQL = sSQL & "FaqA = '"            & DBsafe(iFaqA)   & "', "
    sSQL = sSQL & "FAQCategoryId = "    & sFAQCategoryId  & ", "
    sSQL = sSQL & "publicationstart = " & sPubStartDate   & ", "
    sSQL = sSQL & "publicationend = "   & sPubEndDate     & ", "
    sSQL = sSQL & "faqtype = '"         & UCASE(iFAQType) & "', "
    sSQL = sSQL & "lastupdatedbyid = "  & iUserID         & ", "
    sSQL = sSQL & "lastupdateddate = '" & now()           & "' "
    sSQL = sSQL & " WHERE orgid = " & iOrgID
    sSQL = sSQL & " AND FaqID = "   & iFaqID 

  		set oFaqUpdate = Server.CreateObject("ADODB.Recordset")
	  	oFaqUpdate.Open sSQL, Application("DSN"), 3, 1

    set oFaqUpdate = nothing

    lcl_redirect_url = "manage_faq.asp?ifaqid=" & iFaqID & "&faqtype=" & iFAQType & "&success=SU"

    lcl_success = "SU"

'-- ADD -----------------------------------------------------------------------
 else
'------------------------------------------------------------------------------
  'Check to see if this record is being created from a request.
   if iRequestID <> "" then
      lcl_pushedfrom_requestid = iRequestID
   else
      lcl_pushedfrom_requestid = "NULL"
   end if

		'Want the sequence of the selected category.
 		if CLng(iFAQCategoryId) = CLng(0) then
   			sWhere = "IS NULL "
   else
   			sWhere = "= " & iFAQCategoryId
   end if

 'Set the Created By info
  lcl_createdbyid = iUserID
  lcl_createddate = Now

		'Get the next sequence number 
 		sSQL = "SELECT MAX(sequence) AS maxSeq "
   sSQL = sSQL & " FROM faq "
   sSQL = sSQL & " WHERE OrgID = "     & iOrgID
   sSQL = sSQL & " AND FAQCategoryId " & sWhere
   sSQL = sSQL & " AND UPPER(faqtype) = '" & iFAQType & "'"

 		set oFaq = Server.CreateObject("ADODB.Recordset")
	 	oFaq.Open sSQL, Application("DSN"), 3, 1
		
 		if oFaq.eof then
	   		iSeq = 1
 		else
	   		if isnull(oFaq("maxSeq")) then
     				iSeq = 1
   			else
			     	iSeq = CLng(oFaq("maxSeq")) + CLng(1)
   			end if
 		end if

 		oFAQ.close 
	 	set oFAQ = nothing

		'Insert the new FAQ
 		sSQL = "INSERT INTO FAQ ("
   sSQL = sSQL & "FaqQ, "
   sSQL = sSQL & "FaqA, "
   sSQL = sSQL & "OrgID, "
   sSQL = sSQL & "Sequence, "
   sSQL = sSQL & "FAQCategoryId, "
   sSQL = sSQL & "publicationstart, "
   sSQL = sSQL & "publicationend, "
   sSQL = sSQL & "faqtype, "
   sSQL = sSQL & "createdbyid, "
   sSQL = sSQL & "createddate, "
   sSQL = sSQL & "pushedfrom_requestid "
   sSQL = sSQL & ") VALUES ("
   sSQL = sSQL & "'" & DBsafe(iFaqQ)   & "', "
   sSQL = sSQL & "'" & DBsafe(iFaqA)   & "', "
   sSQL = sSQL &       iOrgID          & ", "
   sSQL = sSQL &       iSeq            & ", "
   sSQL = sSQL &       sFAQCategoryId  & ", "
   sSQL = sSQL &       sPubStartDate   & ", "
   sSQL = sSQL &       sPubEndDate     & ", "
   sSQL = sSQL & "'" & UCASE(iFAQType) & "', "
   sSQL = sSQL &       lcl_createdbyid & ", "
   sSQL = sSQL & "'" & lcl_createddate & "', "
   sSQL = sSQL &       lcl_pushedfrom_requestid
   sSQL = sSQL & ")"

		'Get the FAQid
	 	iFaqID = RunIdentityInsert(sSQL)

  'If this was created from a request then add a record to the activity log
   if iRequestID <> "" then
      intComment   = ""
      lcl_username = getAdminName(lcl_createdbyid)

      sCommentLine = lcl_username & " pushed this request to " & iFAQType
      sCommentLine = sCommentLine & " on " & lcl_createddate

     'Build the Internal Comment
      if intComment <> "" then
         intComment = intComment & "<br />" & sCommentLine
      else
         intComment = sCommentLine
      end if

     'Create a record in the Activity Log
      AddCommentTaskComment trim(intComment), "", "", iRequestID, iUserID, iOrgID, "", "", ""
   end if

   lcl_redirect_url = "list_faq.asp?faqtype=" & iFAQType & "&success=SA"

'------------------------------------------------------------------------------
 end if
'------------------------------------------------------------------------------

'Reorder all faqs in all categories
	ReorderFAQs iFAQCategoryId, iFAQType

'Check to see if there is any aditional processing we will need to do.
 lcl_return_parameters = ""

 if iSendToRSS = "on" then
    lcl_return_parameters = lcl_return_parameters & "&sendTo_RSS=" & iFaqID
 end if

 response.redirect lcl_redirect_url & lcl_return_parameters

end sub

'------------------------------------------------------------------------------
function DBsafe( strDB )
 	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
 	DBsafe = Replace( strDB, "'", "''" )
end function

'------------------------------------------------------------------------------
Function RunIdentityInsert( sInsertStatement )
	Dim sSQL, iReturnValue, oInsert

	iReturnValue = 0

'	response.write "<p>" & sInsertStatement & "</p><br /><br />"
'	response.flush

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSQL = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSQL, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.close
	Set oInsert = Nothing

	RunIdentityInsert = iReturnValue

End Function

'------------------------------------------------------------------------------
sub ReorderFAQs(iFAQCategoryId, p_faqtype)
	Dim sSql, oRs, iFaqCount, iOldCategory

	iFaqCount    = 0
	iOldCategory = -1

	sSQL = "SELECT FaqID, Sequence, ISNULL(FAQCategoryId,0) AS FAQCategoryId "
 sSQL = sSQL & " FROM Faq "
 sSQL = sSQL & " WHERE OrgID = "       & session("orgid")
 sSQL = sSQL & " AND FAQCategoryId = " & iFAQCategoryId
 sSQL = sSQL & " AND UPPER(faqtype) = '" & p_faqtype & "' "
	sSQL = sSQL & " ORDER BY FAQCategoryId, Sequence, FaqID"

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

 if not oRs.eof then
    do while not oRs.eof
 		   	iFaqCount = iFaqCount + 1

     		sSQL = "UPDATE Faq SET Sequence = " & iFaqCount & " WHERE FaqID = " & oRs("FaqID") 
  	   	RunSQL sSql

     		oRs.movenext
    loop
 end if
	
	oRs.close
	set oRs = nothing 

end sub

'------------------------------------------------------------------------------
sub RunSQL( sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 
%>