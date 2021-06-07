<!-- #include file="../includes/common.asp" //-->
<%
'Retrieve all of the parameters
 lcl_dlid                 = request("idlid")
 lcl_name                 = request("sname")
 lcl_description          = request("sdescription")
 lcl_blndisplay           = request("blndisplay")
 lcl_parent_id            = request("sparentid")

 lcl_sc_name              = request("sc_name")
 lcl_sc_description       = request("sc_description")
 lcl_sc_publicly_viewable = request("sc_publicly_viewable")
 lcl_sc_list_type         = request("sc_list_type")
 lcl_sc_orderby           = request("sc_orderby")

'Call subDeletedl(request("idlId"))
call subDeleted1(lcl_dlid, lcl_name, lcl_description, lcl_blndisplay, lcl_parent_id, lcl_sc_name, lcl_sc_description, lcl_sc_publicly_viewable, lcl_sc_list_type, lcl_sc_orderby)
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
' SUB SUBDELETEDL(IDLID)
'--------------------------------------------------------------------------------------------------
'Sub subDeletedl(idlId)
sub subDeleted1(p_dlid, p_name, p_description, p_blndisplay, p_parent_id, p_sc_name, p_sc_description, p_sc_publicly_viewable, p_sc_list_type, p_sc_orderby)

'Delete all jobs/bids assignments related to any sub-categories that are being deleted, if any exist	
 sSQL = "SELECT distributionlistid "
 sSQL = sSQL & " FROM egov_class_distributionlist "
 sSQL = sSQL & " WHERE parentid = " & p_dlid

 set rs = Server.CreateObject("ADODB.Recordset")
 rs.Open sSQL, Application("DSN"), 3, 1

 if not rs.eof then
    while not rs.eof
       sSQL1 = "DELETE FROM egov_distributionlists_jobbids "
       sSQL1 = sSQL1 & " WHERE distributionlistid = " & rs("distributionlistid")

       set rs1 = Server.CreateObject("ADODB.Recordset")
       rs1.Open sSQL1, Application("DSN"), 3, 1

       rs.movenext
    wend
 end if

'Delete any jobs/bids assignments relatd to this category, if any exists
 sSQL2 = "DELETE FROM egov_distributionlists_jobbids "
 sSQL2 = sSQL2 & " WHERE distributionlistid = " & p_dlid

 set rs2 = Server.CreateObject("ADODB.Recordset")
 rs2.Open sSQL2, Application("DSN"), 3, 1

'Delete all sub-categories related to this category, if any exist	
 sSQL3 = "DELETE FROM egov_class_distributionlist "
 sSQL3 = sSQL3 & " WHERE parentid = " & p_dlid

 set rs3 = Server.CreateObject("ADODB.Recordset")
 rs3.Open sSQL3, Application("DSN"), 3, 1

'Now delete the record
	sSQL4 = "DELETE FROM egov_class_distributionlist "
 sSQL4 = sSQL4 & " WHERE distributionlistid = " & p_dlid
	
 set rs4 = Server.CreateObject("ADODB.Recordset")
 rs4.Open sSQL4, Application("DSN"), 3, 1

 lcl_redirect_url = "dl_mgmt.asp?success=SD&sc_name=" & p_sc_name & "&sc_description=" & p_sc_description & "&sc_publicy_viewable=" & p_sc_publicly_viewable & "&sc_list_type=" & p_sc_list_type & "&sc_orderby=" & p_sc_orderby

 response.redirect lcl_redirect_url

End Sub
%>
