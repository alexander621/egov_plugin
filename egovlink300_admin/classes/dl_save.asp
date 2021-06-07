<!-- #include file="../includes/common.asp" //-->
<%
'Retrieve all of the parameters
 lcl_dlid                 = request("idlid")
 lcl_name                 = request("sname")
 lcl_description          = request("sdescription")

 if request("blndisplay") = "1" then
    lcl_blndisplay        = request("blndisplay")
 else
    lcl_blndisplay        = "0"
 end if

 lcl_parent_id            = request("sparentid")

 lcl_sc_name              = request("sc_name")
 lcl_sc_description       = request("sc_description")
 lcl_sc_publicly_viewable = request("sc_publicly_viewable")
 lcl_sc_list_type         = request("sc_list_type")
 lcl_sc_orderby           = request("sc_orderby")

'If the listtype = "BID" then check to see if this record has sub-categories (parentid = lcl_dlid)
'If the category has changed and sub-categories exist then do NOT allow the change of category
 if lcl_sc_list_type = "BID" and lcl_dlid > 0 then
   'Get the current parentid
    sSQL2 = "SELECT parentid "
    sSQL2 = sSQL2 & " FROM egov_class_distributionlist "
    sSQL2 = sSQL2 & " WHERE distributionlistid = " & lcl_dlid

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    rs2.Open sSQL2, Application("DSN"), 3, 1

    if rs2("parentid") = "" OR isnull(rs2("parentid")) then
       lcl_current_parent_id = 0
    else
       lcl_current_parent_id = rs2("parentid")
    end if

    if lcl_parent_id = "" then
       lcl_new_parent_id = 0
    else
       lcl_new_parent_id = lcl_parent_id
    end if

'sSQL1 = "INSERT INTO my_table_dtb (notes) VALUES ('(current: " & lcl_current_parent_id & ") - (new: " & lcl_new_parent_id & ")')"
'Set rs = Server.CreateObject("ADODB.Recordset")
'rs.Open sSQL1, Application("DSN"), 3, 1

    if lcl_new_parent_id > 0 AND lcl_current_parent_id = 0 then
       sSQLc = "SELECT distinct 'Y' AS lcl_exists "
       sSQLc = sSQLc & " FROM egov_class_distributionlist "
       sSQLc = sSQLc & " WHERE parentid = " & lcl_dlid

       Set rsc = Server.CreateObject("ADODB.Recordset")
       rsc.Open sSQLc, Application("DSN"), 3, 1

       if not rsc.eof then
          response.redirect "dl_edit.asp?success=EU&dlid=" & lcl_dlid & "&sc_name=" & lcl_sc_name & "&sc_description=" & lcl_sc_description & "&sc_publicy_viewable=" & lcl_sc_publicly_viewable & "&sc_list_type=" & lcl_sc_list_type & "&sc_orderby=" & lcl_sc_orderby
       end if
    end if
 end if

'Call subSaveDL(request("idlId"), request("sName"), request("sdescription"), request("blndisplay"))
 call subSaveDL(lcl_dlid, lcl_name, lcl_description, lcl_blndisplay, lcl_parent_id, lcl_sc_name, lcl_sc_description, lcl_sc_publicly_viewable, lcl_sc_list_type, lcl_sc_orderby)
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
' SUB SUBSAVEDL(IDLID, SNAME, SDESCRIPTION, BLNDISPLAY)
'--------------------------------------------------------------------------------------------------
'Sub subSaveDL(idlId, sName, sdescription, blndisplay)
sub subSaveDL(p_dlid, p_name, p_description, p_blndisplay, p_parent_id, p_sc_name, p_sc_description, p_sc_publicly_viewable, p_sc_list_type, p_sc_orderby)

'	sName        = DBsafe(p_name)
'	sdescription = DBsafe(p_description)
'	blndisplay   = DBsafe(p_blndisplay)

'	If idlId = "0" Then
 if p_dlid = "0" then
		 'Insert new record
  		sSQL = "INSERT INTO egov_class_distributionlist (OrgID, "
    sSQL = sSQL &                                  " distributionlistname, "
    sSQL = sSQL &                                  " distributionlistdescription, "
    sSQL = sSQL &                                  " distributionlistdisplay, "
    sSQL = sSQL &                                  " distributionlisttype"

    if p_parent_id <> "" then
       sSQL = sSQL & ",parentid"
    end if

    sSQL = sSQL & ") "
    sSQL = sSQL & " VALUES ("
    sSQL = sSQL & Session("OrgID")            & ", "
	  	sSQL = sSQL & "'" & DBsafe(p_name)        & "', "
    sSQL = sSQL & "'" & DBsafe(p_description) & "', "
    sSQL = sSQL & "'" & DBsafe(p_blndisplay)  & "', "
    sSQL = sSQL & "'" & UCASE(p_sc_list_type) & "'"

    if p_parent_id <> "" then
       sSQL = sSQL & "," & p_parent_id
    end if

    sSQL = sSQL & ")"

    lcl_redirect_url = "dl_mgmt.asp?success=SN&sc_name=" & p_sc_name & "&sc_description=" & p_sc_description & "&sc_publicy_viewable=" & p_sc_publicly_viewable & "&sc_list_type=" & p_sc_list_type & "&sc_orderby=" & p_sc_orderby

	Else 
 		'Update existing record
		  sSQL = "UPDATE egov_class_distributionlist SET "
    sSQL = sSQL & " distributionlistname = '"        & DBsafe(p_name)         & "', "
    sSQL = sSQL & " distributionlistdescription = '" & DBsafe(p_description)  & "', "

    if p_blndisplay <> "" then
       sSQL = sSQL & " distributionlistdisplay = '"  & DBsafe(p_blndisplay)   & "', "
    end if

    sSQL = sSQL & " distributionlisttype = '"        & p_sc_list_type & "', "

    if p_parent_id <> "" then
       sSQL = sSQL & " parentid = "                  & p_parent_id    & ""
    else
       sSQL = sSQL & " parentid = NULL "
    end if
		  sSQL = sSQL & " WHERE distributionlistid = " & p_dlid & ""

    lcl_redirect_url = "dl_edit.asp?success=SU&dlid=" & p_dlid & "&sc_name=" & p_sc_name & "&sc_description=" & p_sc_description & "&sc_publicy_viewable=" & p_sc_publicly_viewable & "&sc_list_type=" & p_sc_list_type & "&sc_orderby=" & p_sc_orderby

	End If

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSQL
		.Execute
	End With
	Set oCmd = Nothing

 response.redirect lcl_redirect_url

End Sub
%>