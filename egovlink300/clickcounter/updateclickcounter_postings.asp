<%
 lcl_isAjax     = "N"
 lcl_orgid      = 0
 lcl_userid     = 0
 lcl_posting_id = 0
 lcl_linkID     = "NULL"
 lcl_linkText   = "NULL"
 lcl_linkURL    = "NULL"

 if request("isAjaxRoutine") <> "" then
    lcl_isAjax = UCASE(request("isAjaxRoutine"))
 end if

 if request("orgid") <> "" then
    if isnumeric(request("orgid")) then
       lcl_orgid = request("orgid")
    end if
 end if

 if request("userid") <> "" then
    if isnumeric(request("userid")) then
       lcl_userid = request("userid")
    end if
 end if

 if request("posting_id") <> "" then
    if isnumeric(request("posting_id")) then
       lcl_posting_id = request("posting_id")
    end if
 end if

 if request("linkID") <> "" then
    lcl_linkID = "'" & dbsafe(request("linkID")) & "'"
 end if

 if request("linkText") <> "" then
    lcl_linkText = "'" & dbsafe(request("linkText")) & "'"
 end if

 if request("linkURL") <> "" then
    lcl_linkURL = "'" & dbsafe(request("linkURL")) & "'"
 end if

'If the a link exists then update the counter (insert a record).
' if CLng(lcl_posting_id) > CLng(0) AND lcl_linkID <> "" AND linkTEXT <> "" AND linkURL <> "" then
    sSQL = "INSERT INTO egov_clickcounter_postings ("
    sSQL = sSQL & "orgid, "
    sSQL = sSQL & "userid, "
    sSQL = sSQL & "posting_id, "
    sSQL = sSQL & "clicked_linkid, "
    sSQL = sSQL & "clicked_linktext, "
    sSQL = sSQL & "clicked_linkurl, "
    sSQL = sSQL & "clicked_date"
    sSQL = sSQL & ") VALUES ("
    sSQL = sSQL & lcl_orgid      & ", "
    sSQL = sSQL & lcl_userid     & ", "
    sSQL = sSQL & lcl_posting_id & ", "
    sSQL = sSQL & lcl_linkID     & ", "
    sSQL = sSQL & lcl_linkText   & ", "
    sSQL = sSQL & lcl_linkURL    & ", "
    sSQL = sSQL & "'" & now()    & "' "
    sSQL = sSQL & ")"

   	set oUpdateClickCount = Server.CreateObject("ADODB.Recordset")
  	 oUpdateClickCount.Open sSQL, Application("DSN"), 3, 1

    set oUpdateClickCount = nothing
' else
'    if lcl_isAjax = "Y" then
'       response.write "Failed to update 'Click Counter' - Error in AJAX Routine"
'    else
'       response.write "../postings_info.asp?success=AJAX_ERROR"
'    end if
' end if

'response.write "Updated Counter!"

'------------------------------------------------------------------------------
function dbsafe(p_value)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = replace(p_value,"'","''")
  end if

  dbsafe = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "') "
 	set oDTB = Server.CreateObject("ADODB.Recordset")
	 oDTB.Open sSQL, Application("DSN"), 3, 1

  set oDTB = nothing

end sub
%>