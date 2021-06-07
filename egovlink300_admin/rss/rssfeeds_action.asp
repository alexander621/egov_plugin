<!-- #include file="../includes/common.asp" //-->
<%
Call updateRSSFeed(request("user_action"),request("feedid"), request("feedname"), request("rsstitle"), request("feedurl"), request("description"), request("isActive"), request("feature"), request("orgtitle"))

'------------------------------------------------------------------------------
sub updateRSSFeed(iAction, ifeedid, ifeedname, irsstitle, ifeedurl, idescription, p_isActive, ifeature, iorgtitle)

 if ifeedid <> "" then
    sfeedid = CLng(ifeedid)
 else
    sfeedid = 0
 end if

 if ifeedname = "" then
    sfeedname = "NULL"
 else
    sfeedname = "'" & dbsafe(ifeedname) & "'"
 end if

 if p_isActive = "on" then
  		sisActive = 1
 else
		  sisActive = 0
 end if

 if irsstitle = "" then
  		srsstitle = "NULL"
 else
  		srsstitle = "'" & dbsafe(irsstitle) & "'"
 end if

 if idescription = "" then
  		sdescription = "NULL"
 else
  		sdescription = "'" & dbsafe(LEFT(idescription,8000)) & "'"
 end if

 if ifeedurl = "" then
  		sfeedurl = "NULL"
 else
  		sfeedurl = "'" & dbsafe(ifeedurl) & "'"
 end if

 if ifeature = "" then
    sfeature = "NULL"
 else
    sfeature = "'" & dbsafe(ifeature) & "'"
 end if

'The rss feed exists, so update it
 if iAction = "UPDATE" then
  		sSQL = "UPDATE egov_rssfeeds SET "
    sSQL = sSQL & "feedname = "    & sfeedname    & ", "
    sSQL = sSQL & "isActive = "    & sisActive    & ", "
    sSQL = sSQL & "title = "       & srsstitle    & ", "
    sSQL = sSQL & "description = " & sdescription & ", "
    sSQL = sSQL & "feedurl = "     & sfeedurl     & ", "
    sSQL = sSQL & "feature = "     & sfeature
    sSQL = sSQL & " WHERE feedid = " & sfeedid

  		set oRSSFeedUpdate = Server.CreateObject("ADODB.Recordset")
	  	oRSSFeedUpdate.Open sSQL, Application("DSN"), 3, 1

    set oRSSFeedUpdate = nothing

    lcl_redirect_url = "rssfeeds_maint.asp?feedid=" & sfeedid & "&success=SU"

   'Maintain the OrgTitle on egov_rssfeeds_orgtitles.
    deleteRSSOrgTitle session("orgid"), sfeedid

    if iorgtitle <> "" then
       insertRSSOrgTitle session("orgid"), sfeedid, iorgtitle
    end if

'------------------------------------------------------------------------------
 elseif iAction = "DELETE" then
'------------------------------------------------------------------------------

   'BEGIN: Delete all custom OrgTitles associated to this feed. ---------------
    sSQL = "DELETE FROM egov_rssfeeds_orgtitles "
    sSQL = sSQL & " WHERE feedid = "  & sfeedid

  	 set oDeleteRSSOrgTitles = Server.CreateObject("ADODB.Recordset")
 	  oDeleteRSSOrgTitles.Open sSQL, Application("DSN"), 3, 1
   'END: Delete all custom OrgTitles associated to this feed. -----------------

   'BEGIN: Delete all of the RSS Items associated to this RSS Feed. -----------
    sSQL = "DELETE FROM egov_rss "
    sSQL = sSQL & " WHERE orgid = " & session("orgid")
    sSQL = sSQL & " AND feedid = "  & sfeedid

  	 set oDeleteRSSItems = Server.CreateObject("ADODB.Recordset")
 	  oDeleteRSSItems.Open sSQL, Application("DSN"), 3, 1
   'END: Delete all of the RSS Items associated to this RSS Feed. -------------

   'BEGIN: Delete the RSS Feed. -----------------------------------------------
    sSQL = "DELETE FROM egov_rssfeeds "
    sSQL = sSQL & " WHERE feedid = "  & sfeedid

  	 set oDeleteRSSFeed = Server.CreateObject("ADODB.Recordset")
 	  oDeleteRSSFeed.Open sSQL, Application("DSN"), 3, 1
   'END: Delete the RSS Feed. -------------------------------------------------

    set oDeleteRSSOrgTitles = nothing
    set oDeleteRSSItems     = nothing
    set oDeleteRSSFeed      = nothing

    lcl_redirect_url = "rssfeeds_list.asp?success=SD"

'------------------------------------------------------------------------------
 else  'New RSS Feed
'------------------------------------------------------------------------------

 		'Insert the new RSS Feed
  		sSQL = "INSERT INTO egov_rssfeeds ("
    sSQL = sSQL & "feedname, "
    sSQL = sSQL & "isActive, "
    sSQL = sSQL & "title, "
    sSQL = sSQL & "description,"
    sSQL = sSQL & "feedurl, "
    sSQL = sSQL & "lastbuilddate, "
    sSQL = sSQL & "feature "
    sSQL = sSQL & ") VALUES ("
    sSQL = sSQL & sfeedname        & ", "
    sSQL = sSQL & sisActive        & ", "
    sSQL = sSQL & srsstitle        & ", "
    sSQL = sSQL & sdescription     & ", "
    sSQL = sSQL & sfeedurl         & ", "
    sSQL = sSQL & "NULL, "
    sSQL = sSQL & sfeature
    sSQL = sSQL & ")"

    lcl_redirect_url = "rssfeeds_maint.asp?success=SA"

    if iAction = "ADD" then
    		'Get the feedid
   	  	lcl_feedid = RunIdentityInsert(sSQL)

       lcl_redirect_url = lcl_redirect_url & "&feedid=" & lcl_feedid
    end if

   'Maintain the OrgTitle on egov_rssfeeds_orgtitles.
    deleteRSSOrgTitle session("orgid"), lcl_feedid

    if iorgtitle <> "" then
       insertRSSOrgTitle session("orgid"), lcl_feedid, iorgtitle
    end if

 end if

 response.redirect lcl_redirect_url

end sub

'------------------------------------------------------------------------------
function DBsafe( strDB )
 	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
 	DBsafe = Replace( strDB, "'", "''" )
end function

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
sub deleteRSSOrgTitle(p_orgid,p_feedid)

  if p_orgid <> "" AND p_feedid <> "" then
     sSQL = "DELETE FROM egov_rssfeeds_orgtitles "
     sSQL = sSQL & " WHERE orgid = " & p_orgid
     sSQL = sSQL & " AND feedid = "  & p_feedid

     set oRSSOrgTitleDel = Server.CreateObject("ADODB.Recordset")
     oRSSOrgTitleDel.Open sSQL, Application("DSN"), 3, 1

     set oRSSOrgTitleDel = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub insertRSSOrgTitle(p_orgid, p_feedid, p_orgtitle)

  if p_orgid <> "" AND p_feedid <> "" AND p_orgtitle <> "" then
     sSQL = "INSERT INTO egov_rssfeeds_orgtitles(orgid,feedid,orgtitle) VALUES ("
     sSQL = sSQL &       p_orgid            & ", "
     sSQL = sSQL &       p_feedid           & ", "
     sSQL = sSQL & "'" & dbsafe(p_orgtitle) & "' "
     sSQL = sSQL & ")"

     set oRSSOrgTitleIns = Server.CreateObject("ADODB.Recordset")
     oRSSOrgTitleIns.Open sSQL, Application("DSN"), 3, 1

     set oRSSOrgTitleIns = nothing
  end if

end sub
%>