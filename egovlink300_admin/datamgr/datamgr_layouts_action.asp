<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
  if request("user_action") = "" then
     response.redirect "datamgr_layouts_list.asp"
  end if

 'Setup variables
  lcl_useraction        = ""
  lcl_layoutid          = 0
  lcl_orgid             = request("orgid")
  sLayoutName           = request("layoutname")
  lcl_isActive          = 1
  lcl_useLayoutSections = 1
  sTotalColumns         = "1"
  sColumnWidthLeft      = "100"
  sColumnWidthMiddle    = "0"
  sColumnWidthRight     = "0"
  lcl_userid            = session("userid")
  lcl_current_date      = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"
  lcl_redirect_url      = "datamgr_list.asp"

  if request("user_action") <> "" then
     lcl_useraction = UCASE(request("user_action"))
  end if

  if request("layoutid") <> "" then
     lcl_layoutid = request("layoutid")
  end if

 'Build return parameters
  lcl_url_parameters = ""

 'Execute the user's action
  if lcl_useraction = "DELETE" then
    'Delete the Map-Point
     sSQL = "DELETE FROM egov_dm_layouts WHERE layoutid = " & lcl_layoutid

   		set oDeleteDMLayout = Server.CreateObject("ADODB.Recordset")
    	oDeleteDMLayout.Open sSQL, Application("DSN"), 3, 1

     set oDeleteDMLayout = nothing

     lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SD")
     lcl_redirect_url   = "datamgr_layouts_list.asp" & lcl_url_parameters

  else
    'BEGIN: Format the columns for the table ----------------------------------
     sLayoutName = formatFieldforInsertUpdate(sLayoutName)

     if request("isActive") = "Y" then
        lcl_isActive = 1
     else
        lcl_isActive = 0
     end if

     if request("useLayoutSections") = "Y" then
        lcl_useLayoutSections = 1
     else
        lcl_useLayoutSections = 0
     end if

     if request("totalcolumns") <> "" then
        sTotalColumns = request("totalcolumns")
     end if

     if request("columnwidth_left") <> "" then
        sColumnWidthLeft = request("columnwidth_left")
     end if

     if request("columnwidth_middle") <> "" then
        sColumnWidthMiddle = request("columnwidth_middle")
     end if

     if request("columnwidth_right") <> "" then
        sColumnWidthRight = request("columnwidth_right")
     end if
    'END: Format the columns for the table ------------------------------------

     if lcl_useraction = "UPDATE" then

      		sSQL = "UPDATE egov_dm_layouts SET "
        sSQL = sSQL & "layoutname = "         & sLayoutName           & ", "
        sSQL = sSQL & "isActive = "           & lcl_isActive          & ", "
        sSQL = sSQL & "useLayoutSections = "  & lcl_useLayoutSections & ", "
        sSQL = sSQL & "totalcolumns = "       & sTotalColumns         & ", "
        sSQL = sSQL & "columnwidth_left = "   & sColumnWidthLeft      & ", "
        sSQL = sSQL & "columnwidth_middle = " & sColumnWidthMiddle    & ", "
        sSQL = sSQL & "columnwidth_right = "  & sColumnWidthRight     & ", "
        sSQL = sSQL & "lastmodifiedbyid = "   & lcl_userid            & ", "
        sSQL = sSQL & "lastmodifiedbydate = " & lcl_current_date
        sSQL = sSQL & " WHERE layoutid = " & lcl_layoutid

      		set oUpdateDMLayout = Server.CreateObject("ADODB.Recordset")
	      	oUpdateDMLayout.Open sSQL, Application("DSN"), 3, 1

        set oUpdateDMLayout = nothing

        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "layoutid", lcl_layoutid)
        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SU")
        lcl_redirect_url   = "datamgr_layouts_maint.asp" & lcl_url_parameters

    '---------------------------------------------------------------------------
     else  'New DataMgr Layout
    '---------------------------------------------------------------------------
        sCreatedByID   = lcl_userid
        sCreatedByDate = lcl_current_date

     		'Insert the new Map-Point
   	   	sSQL = "INSERT INTO egov_dm_layouts ("
        sSQL = sSQL & "layoutname, "
        sSQL = sSQL & "isActive, "
        sSQL = sSQL & "useLayoutSections, "
        sSQL = sSQL & "totalcolumns, "
        sSQL = sSQL & "columnwidth_left, "
        sSQL = sSQL & "columnwidth_middle, "
        sSQL = sSQL & "columnwidth_right, "
        sSQL = sSQL & "createdbyid, "
        sSQL = sSQL & "createdbydate, "
        sSQL = sSQL & "lastmodifiedbyid, "
        sSQL = sSQL & "lastmodifiedbydate "
        sSQL = sSQL & ") VALUES ("
        sSQL = sSQL & sLayoutName           & ", "
        sSQL = sSQL & lcl_isActive          & ", "
        sSQL = sSQL & lcl_useLayoutSections & ", "
        sSQL = sSQL & sTotalColumns         & ", "
        sSQL = sSQL & sColumnWidthLeft      & ", "
        sSQL = sSQL & sColumnWidthMiddle    & ", "
        sSQL = sSQL & sColumnWidthRight     & ", "
        sSQL = sSQL & sCreatedByID          & ", "
        sSQL = sSQL & sCreatedByDate        & ", "
        sSQL = sSQL & "NULL,NULL"
        sSQL = sSQL & ")"

     		'Get the MapPointID
    	  	lcl_layoutid = RunIdentityInsert(sSQL)

        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "layoutid", lcl_layoutid)
        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SA")
        lcl_redirect_url   = "datamgr_layouts_maint.asp" & lcl_url_parameters
     end if
  end if

  response.redirect lcl_redirect_url
%>