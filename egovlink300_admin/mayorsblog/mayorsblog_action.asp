<!-- #include file="../includes/common.asp" //-->
<%
Call updateMayorsBlog(request("user_action"),request("blogid"), request("userid"), request("blogtitle"), request("article"), _
                      request("isInactive"))

'------------------------------------------------------------------------------
sub updateMayorsBlog(iAction, iBlogID, iUserID, iBlogTitle, iArticle, iIsInactive)

 if iBlogID <> "" then
    sBlogID = CLng(iBlogID)
 else
    sBlogID = 0
 end if

 if iUserID <> "" then
    sUserID = CLng(iUserID)
 else
    sUserID = 0
 end if

 if iBlogTitle = "" then
  		sBlogTitle = "NULL"
 else
  		sBlogTitle = "'" & dbsafe(iBlogTitle) & "'"
 end if

 if iArticle = "" then
  		sArticle = "NULL"
 else
  		sArticle = "'" & dbsafe(iArticle) & "'"
 end if

 if iIsInactive = "on" then
  		sIsInactive = 0
 else
		  sIsInactive = 1
 end if

'The blog exists, so update it
 if iAction = "UPDATE" then
  		sSQL = "UPDATE egov_mayorsblog SET "
    sSQL = sSQL & "userid = "              & sUserID           & ", "
    sSQL = sSQL & "title = "               & sBlogTitle        & ", "
    sSQL = sSQL & "article = "             & sArticle          & ", "
    sSQL = sSQL & "isInactive = "          & sIsInactive       & ", "
    sSQL = sSQL & "lastmodifiedbyid = "    & session("userid") & ", "
    sSQL = sSQL & "lastmodifiedbydate = '" & dbsafe(ConvertDateTimetoTimeZone()) & "' "
    sSQL = sSQL & " WHERE blogid = " & sBlogID

  		set oBlogUpdate = Server.CreateObject("ADODB.Recordset")
	  	oBlogUpdate.Open sSQL, Application("DSN"), 3, 1

    set oBlogUpdate = nothing

    lcl_redirect_url = "mayorsblog_maint.asp?blogid=" & sBlogID & "&success=SU"

'------------------------------------------------------------------------------
 else  'New Blog
'------------------------------------------------------------------------------
    sCreatedByID   = session("userid")
    sCreatedByDate = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"

 		'Insert the new Blog
  		sSQL = "INSERT INTO egov_mayorsblog ("
    sSQL = sSQL & "orgid, "
    sSQL = sSQL & "userid, "
    sSQL = sSQL & "title, "
    sSQL = sSQL & "article, "
    sSQL = sSQL & "createdbyid, "
    sSQL = sSQL & "createdbydate, "
    sSQL = sSQL & "isInactive, "
    sSQL = sSQL & "lastmodifiedbyid,"
    sSQL = sSQL & "lastmodifiedbydate"
    sSQL = sSQL & ") VALUES ("
    sSQL = sSQL & session("orgid") & ", "
    sSQL = sSQL & sUserID          & ", "
    sSQL = sSQL & sBlogTitle       & ", "
    sSQL = sSQL & sArticle         & ", "
    sSQL = sSQL & sCreatedByID     & ", "
    sSQL = sSQL & sCreatedByDate   & ", "
    sSQL = sSQL & sIsInactive      & ", "
    sSQL = sSQL & "NULL,NULL"
    sSQL = sSQL & ")"

 		'Get the BlogID
	  	sBlogID = RunIdentityInsert(sSQL)

    lcl_redirect_url = "mayorsblog_maint.asp?success=SA"

    if iAction = "ADD" then
       lcl_redirect_url = lcl_redirect_url & "&blogid=" & sBlogID
    end if

 end if

'Check to see if there is any aditional processing we will need to do.
 lcl_return_parameters = ""

 if request("sendTo_RSS") = "on" then
    lcl_return_parameters = lcl_return_parameters & "&sendTo_RSS=" & sBlogID
 end if

 response.redirect lcl_redirect_url & lcl_return_parameters

end sub

'------------------------------------------------------------------------------
function RunIdentityInsert( sInsertStatement )
	 Dim sSQL, iReturnValue, oInsert

	 iReturnValue = 0

	'Insert new row into database and get rowid
 	sSQL = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

 	set oInsert = Server.CreateObject("ADODB.Recordset")
	 oInsert.Open sSQL, Application("DSN"), 3, 3

 	iReturnValue = oInsert("ROWID")

 	oInsert.close
	 set oInsert = nothing

 	RunIdentityInsert = iReturnValue

end function

'------------------------------------------------------------------------------
function dbsafe(p_value)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = p_value
     lcl_return = replace(lcl_return,"'","''")
  end if

  dbsafe = lcl_return

end function
%>