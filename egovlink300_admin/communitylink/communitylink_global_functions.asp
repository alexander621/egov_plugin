<%
'Check for org features
 lcl_orghasfeature_registration       = orghasfeature("registration")
 lcl_orghasfeature_administrationlink = orghasfeature("AdministrationLink")
 lcl_orghasfeature_payments           = orghasfeature("payments")
 lcl_orghasfeature_action_line        = orghasfeature("action line")
 lcl_orghasfeature_activities         = orghasfeature("activities")
 lcl_orghasfeature_facilities         = orghasfeature("facilities")
 lcl_orghasfeature_memberships        = orghasfeature("memberships")
 lcl_orghasfeature_gifts              = orghasfeature("gifts")
 lcl_orghasfeature_bid_postings       = orghasfeature("bid_postings")

'Check for PublicCanViewFeature
 lcl_publiccanviewfeature_payments     = publiccanviewfeature(session("orgid"),"payments")
 lcl_publiccanviewfeature_action_line  = publiccanviewfeature(session("orgid"),"action line")
 lcl_publiccanviewfeature_activities   = publiccanviewfeature(session("orgid"),"activities")
 lcl_publiccanviewfeature_facilities   = publiccanviewfeature(session("orgid"),"facilities")
 lcl_publiccanviewfeature_memberships  = publiccanviewfeature(session("orgid"),"memberships")
 lcl_publiccanviewfeature_gifts        = publiccanviewfeature(session("orgid"),"gifts")
 lcl_publiccanviewfeature_bid_postings = publiccanviewfeature(session("orgid"),"big postings")

'------------------------------------------------------------------------------
sub displaySideMenubar(p_orgid, iSideMenuOptionBGColor, iSideMenuOptionBGColorHover, iSideMenuOptionAlignment, p_isEgovHomePage)

  if p_isEgovHomePage <> "" then
     lcl_isEgovHomePage = p_isEgovHomePage
  else
     lcl_isEgovHomePage = 0
  end if

 'Display the City Home and E-Gov Home URLs
  displaySideMenubarOption "0", iSideMenuOptionAlignment, "#communitylink_preview", "City Home"
  displaySideMenubarOption "1", iSideMenuOptionAlignment, "#communitylink_preview", "E-Gov Home"

		sSQL = "SELECT O.OrgEgovWebsiteURL, isnull(FO.publicurl,F.publicURL) as publicURL, "
		sSQL = sSQL & "isnull(FO.featurename,F.featurename) as featurename, f.feature "
		sSQL = sSQL & " FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSQL = sSQL & " WHERE FO.publiccanview = 1 "
		sSQL = sSQL & " AND F.haspublicview = 1  "
		sSQL = sSQL & " AND O.orgid = FO.orgid  "
		sSQL = sSQL & " AND FO.featureid = F.featureid  "
		sSQL = sSQL & " AND O.orgid = " & p_orgid
		sSQL = sSQL & " ORDER BY FO.publicdisplayorder, F.publicdisplayorder "

		set oSideNav = Server.CreateObject("ADODB.Recordset")
		oSideNav.Open sSQL, Application("DSN"), 3, 1

  i = 1

		do while not oSideNav.eof

    'Only display the menu option if:
    '1. the option (feature) does NOT equal "communitylink".
    '2. the option (featuer) DOES equal "communitylink" AND it has NOT been set to be the "e-gov home" page.
     if UCASE(oSideNav("feature")) <> "COMMUNITYLINK" OR (UCASE(oSideNav("feature")) = "COMMUNITYLINK" and NOT lcl_isEgovHomePage) then
        i = i + 1

        displaySideMenubarOption i, iSideMenuOptionAlignment, "#communitylink_preview", oSideNav("featurename")
     end if

     oSideNav.movenext
  loop

		oSideNav.close
		set oSideNav = nothing

	'Add the login link for those that have this
		if lcl_orghasfeature_registration then
     i = i + 1

     displaySideMenubarOption i, iSideMenuOptionAlignment, "#communitylink_preview", "Login"

  end if

end sub

'------------------------------------------------------------------------------
sub displaySideMenubarOption(pID, p_alignment, p_url, p_label)

  lcl_onmouseover = " onmouseover=""setupMenuOption('OVER','" & pID & "');"""
  lcl_onmouseout  = " onmouseout=""setupMenuOption('OUT','"   & pID & "');"""
  lcl_onclick     = " onclick=""location.href='" & p_URL & "';"""

  response.write "<div id=""sideMenuBar" & pID & """ class=""sideMenuBar"" align=""" & p_alignment & """" & lcl_onmouseover & lcl_onmouseout & lcl_onclick & ">" & vbcrlf
  response.write "<a href=""" & p_URL & """ id=""sideMenuBarOption" & pID & """ class=""sideMenuBarOption"">" & p_label & "</a>" & vbcrlf
  response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub ShowPublicDefaultFooterNav(p_orgid, iCount, p_isEgovHomePage)
		Dim sSQL, oNav, sNav, iTotalCount

  if p_isEgovHomePage <> "" then
     lcl_isEgovHomePage = p_isEgovHomePage
  else
     lcl_isEgovHomePage = 0
  end if

		sSQL = "SELECT O.OrgEgovWebsiteURL, isnull(FO.publicurl,F.publicURL) as publicURL, "
		sSQL = sSQL & "isnull(FO.featurename,F.featurename) as featurename, f.feature "
		sSQL = sSQL & " FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSQL = sSQL & " WHERE FO.publiccanview = 1 "
		sSQL = sSQL & " AND F.haspublicview = 1 "
		sSQL = sSQL & " AND O.orgid = FO.orgid "
		sSQL = sSQL & " AND FO.featureid = F.featureid "
		sSQL = sSQL & " AND O.orgid = " & p_orgid
		sSQL = sSQL & " ORDER BY FO.publicdisplayorder, F.publicdisplayorder"

		set oFooter = Server.CreateObject("ADODB.Recordset")
		oFooter.Open sSQL, Application("DSN"), 3, 1

		iTotalCount = oFooter.recordcount

  do while not oFooter.eof

    'Only display the footer option if:
    '1. the option (feature) does NOT equal "communitylink".
    '2. the option (featuer) DOES equal "communitylink" AND it has NOT been set to be the "e-gov home" page.
     if UCASE(oFooter("feature")) <> "COMMUNITYLINK" OR (UCASE(oFooter("feature")) = "COMMUNITYLINK" and NOT lcl_isEgovHomePage) then
        if (iCount Mod 6) = 0 then
       				response.write "<br />"
     			else 
       				response.write " | "
        end if


     			response.write "<a href=""#communitylink_preview"" class=""footerOption"">" & oFooter("featurename") & "</a>" & vbcrlf

     			iCount = iCount + 1
     end if

  			oFooter.movenext
  loop

		oFooter.close
		set oFooter = nothing

end sub

'------------------------------------------------------------------------------
function getOrgTagLine(p_orgid)

  lcl_return = ""

  if p_orgid <> "" then
     sSQL = "SELECT orgtagline "
     sSQL = sSQL & " FROM organizations "
     sSQL = sSQL & " WHERE orgid = " & p_orgid

   		set oTagLine = Server.CreateObject("ADODB.Recordset")
   		oTagLine.Open sSQL, Application("DSN"), 3, 1

     if not oTagLine.eof then
        lcl_return = oTagLine("orgtagline")
     end if

     oTagLine.close
     set oTagLine = nothing

  end if

  getOrgTagLine = lcl_return

end function

'------------------------------------------------------------------------------
sub displayCommunityLinkOptions(iOptionName, iCurrentValue)

  sSQL = "SELECT cl_optionid, optionvalue, optionlabel, isDefault "
  sSQL = sSQL & " FROM egov_communitylink_options "
  sSQL = sSQL & " WHERE isActive = 1 "
  sSQL = sSQL & " AND UPPER(optionname) = '" & UCASE(iOptionName) & "' "
  sSQL = sSQL & " ORDER BY displayorder "

		set oCLOptions = Server.CreateObject("ADODB.Recordset")
		oCLOptions.Open sSQL, Application("DSN"), 3, 1

  if not oCLOptions.eof then
     do while not oCLOptions.eof

       'Determine if a value is selected.  If no value has been saved then look for a default.
        lcl_selected_cloption = ""

        if iCurrentValue <> "" then
           if UCASE(iCurrentValue) = UCASE(oCLOptions("optionvalue")) then
              lcl_selected_cloption = " selected=""selected"""
           end if
        else
           if oCLOptions("isDefault") then
              lcl_selected_cloption = " selected=""selected"""
           end if
        end if

        response.write "  <option value=""" & oCLOptions("optionvalue") & """" & lcl_selected_cloption & ">" & oCLOptions("optionlabel") & "</option>" & vbcrlf

        oCLOptions.movenext
     loop

  end if

  oCLOptions.close
  set oCLOptions = nothing

end sub

'------------------------------------------------------------------------------
sub displayRSSFeedOptions(iCurrentValue)

  response.write "  <option value=""0"">&nbsp;</option>" & vbclrf

  sSQL = "SELECT feedid, feedname "
  sSQL = sSQL & " FROM egov_rssfeeds "
  sSQL = sSQL & " WHERE isActive = 1 "
  sSQL = sSQL & " ORDER BY feedname "

		set oRSSFeedOptions = Server.CreateObject("ADODB.Recordset")
		oRSSFeedOptions.Open sSQL, Application("DSN"), 3, 1

  if not oRSSFeedOptions.eof then
     do while not oRSSFeedOptions.eof

       'Determine if a value is selected.  If no value has been saved then look for a default.
        if iCurrentValue <> "" then
           if UCASE(iCurrentValue) = UCASE(oRSSFeedOptions("feedid")) then
              lcl_selected_option = " selected=""selected"""
           else
              lcl_selected_option = ""
           end if
        end if

        response.write "  <option value=""" & oRSSFeedOptions("feedid") & """" & lcl_selected_option & ">" & oRSSFeedOptions("feedname") & "</option>" & vbcrlf

        oRSSFeedOptions.movenext
     loop

  end if

  oRSSFeedOptions.close
  set oRSSFeedOptions = nothing

end sub

'------------------------------------------------------------------------------
function getCLOptionDefault(iOptionName)

  lcl_return = ""

  if iOptionName <> "" then
     sSQL = "SELECT optionvalue "
     sSQL = sSQL & " FROM egov_communitylink_options "
     sSQL = sSQL & " WHERE isActive = 1 "
     sSQL = sSQL & " AND isDefault = 1 "
     sSQL = sSQL & " AND UPPER(optionname) = '" & UCASE(iOptionName) & "' "

   		set oCLOptionDefault = Server.CreateObject("ADODB.Recordset")
   		oCLOptionDefault.Open sSQL, Application("DSN"), 3, 1

     if not oCLOptionDefault.eof then
        lcl_return = oCLOptionDefault("optionvalue")
     end if

     oCLOptionDefault.close
     set oCLOptionDefault = nothing

  end if

  getCLOptionDefault = lcl_return

end function

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SR" then
        lcl_return = "Successfully Reordered..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
function getCommunityLinkID(p_orgid, p_userid)
  lcl_return = 0

  if p_orgid <> "" then
     if isnumeric(p_orgid) then
       'First check to see if the org has a community link record.
        sSQL = "SELECT communitylinkid "
        sSQL = sSQL & " FROM egov_communitylink "
        sSQL = sSQL & " WHERE orgid = " & p_orgid

        set oCLExists = Server.CreateObject("ADODB.Recordset")
       	oCLExists.Open sSQL, Application("DSN"), 3, 1

        if not oCLExists.eof then
           lcl_return = oCLExists("communitylinkid")
        end if

        oCLExists.close
        set oCLExists = nothing
     end if
  end if

  if Clng(lcl_return) = Clng(0) then
     lcl_return = createCommunityLink(p_orgid, p_userid)
  end if

  getCommunityLinkID = lcl_return

end function

'------------------------------------------------------------------------------
function createCommunityLink(p_orgid, p_userid)
  lcl_return = 0

  if p_orgid <> "" then
    'First check to see if the org has a community link record.
     sSQL = "INSERT INTO egov_communitylink (orgid, lastmodifiedbyid, lastmodifiedbydate) VALUES ("
     sSQL = sSQL &       p_orgid                             & ", "
     sSQL = sSQL &       p_userid                            & ", "
     sSQL = sSQL & "'" & dbsafe(ConvertDateTimetoTimeZone()) & "' "
     sSQL = sSQL & ") "

     lcl_return = runIdentityInsert(sSQL)

  end if

  createCommunityLink = lcl_return

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
sub getCommunityLinkInfo(ByVal iCommunityLinkID, ByVal p_orgid, ByRef lcl_isEgovHomePage, ByRef lcl_website_size, _
                         ByRef lcl_website_size_customsize, ByRef lcl_website_alignment, ByRef lcl_website_bgcolor, _
                         ByRef lcl_showlogo, ByRef lcl_logo_filename, ByRef lcl_logo_filenamebg, ByRef lcl_logo_alignment, _
                         ByRef lcl_showtopbar, ByRef lcl_topbar_bgcolor, ByRef lcl_topbar_fonttype, ByRef lcl_topbar_fontcolor, _
                         ByRef lcl_topbar_fontcolorhover, ByRef lcl_showsidemenubar, ByRef lcl_sidemenubar_alignment, _
                         ByRef lcl_sidemenuoption_bgcolor, ByRef lcl_sidemenuoption_bgcolorhover, ByRef lcl_sidemenuoption_alignment, _
                         ByRef lcl_sidemenuoption_fonttype, ByRef lcl_sidemenuoption_fontcolor, ByRef lcl_sidemenuoption_fontcolorhover, _
                         ByRef lcl_showpageheader, ByRef lcl_pageheader_alignment, ByRef lcl_pageheader_fontsize, _
                         ByRef lcl_pageheader_fontcolor, ByRef lcl_pageheader_fonttype, ByRef lcl_pageheader_bgcolor, _
                         ByRef lcl_showfooter, ByRef lcl_footer_bgcolor, ByRef lcl_footer_fonttype, ByRef lcl_footer_fontcolor, _
                         ByRef lcl_footer_fontcolorhover, ByRef lcl_showRSS, ByRef lcl_url_twitter, ByRef lcl_url_facebook, _
                         ByRef lcl_url_myspace, ByRef lcl_url_blogger )

 'Setup defaults
  lcl_isEgovHomePage                = 0
  lcl_website_size                  = getCLOptionDefault("WEBSITE_SIZE")
  lcl_website_size_customsize       = ""
  lcl_website_alignment             = getCLOptionDefault("WEBSITE_ALIGN")
  lcl_website_bgcolor               = "ffffff"
  lcl_showlogo                      = 1
  lcl_logo_filename                 = ""
  lcl_logo_filenamebg               = ""
  lcl_logo_alignment                = getCLOptionDefault("WEBSITE_LOGO_ALIGN")
  lcl_showtopbar                    = 1
  lcl_topbar_fonttype               = getCLOptionDefault("TOPBAR_FONTTYPE")
  lcl_topbar_bgcolor                = "ffffff"
  lcl_topbar_fontcolor              = "000000"
  lcl_topbar_fontcolorhover         = "000000"
  lcl_showsidemenubar               = 1
  lcl_sidemenubar_alignment         = getCLOptionDefault("SIDEMENUBAR_ALIGN")
  lcl_sidemenuoption_bgcolor        = "efefef"
  lcl_sidemenuoption_bgcolorhover   = "c0c0c0"
  lcl_sidemenuoption_alignment      = getCLOptionDefault("SIDEMENUOPT_TEXTALIGN")
  lcl_sidemenuoption_fonttype       = getCLOptionDefault("SIDEMENUOPT_FONTTYPE")
  lcl_sidemenuoption_fontcolor      = "000000"
  lcl_sidemenuoption_fontcolorhover = "000000"
  lcl_showpageheader                = 1
  lcl_pageheader_alignment          = getCLOptionDefault("PAGEHEADER_ALIGN")
  lcl_pageheader_fontsize           = "12"
  lcl_pageheader_fontcolor          = "000000"
  lcl_pageheader_fonttype           = getCLOptionDefault("PAGEHEADER_FONTTYPE")
  lcl_pageheader_bgcolor            = "efefef"
  lcl_showRSS                       = 1
  lcl_showfooter                    = 1
  lcl_footer_fonttype               = getCLOptionDefault("FOOTER_FONTTYPE")
  lcl_footer_bgcolor                = "ffffff"
  lcl_footer_fontcolor              = "000000"
  lcl_footer_fontcolorhover         = "000000"
  lcl_url_twitter                   = ""
  lcl_url_facebook                  = ""
  lcl_url_myspace                   = ""
  lcl_url_blogger                   = ""

 'Pull data record exists.
  if iCommunityLinkID <> "" then
     sSQL = "SELECT "
     sSQL = sSQL & " isEgovHomePage, "
     sSQL = sSQL & " isnull(website_size,'"                  & lcl_website_size                  & "') AS website_size, "
     sSQL = sSQL & " website_size_customsize, "
     sSQL = sSQL & " isnull(website_alignment,'"             & lcl_website_alignment             & "') AS website_alignment, "
     sSQL = sSQL & " isnull(website_bgcolor,'"               & lcl_website_bgcolor               & "') AS website_bgcolor, "
     sSQL = sSQL & " isnull(showlogo,'"                      & lcl_showlogo                      & "') AS showlogo, "
     sSQL = sSQL & " logo_filename, "
     sSQL = sSQL & " logo_filenamebg, "
     sSQL = sSQL & " isnull(logo_alignment,'"                & lcl_logo_alignment                & "') AS logo_alignment, "
     sSQL = sSQL & " isnull(showtopbar,'"                    & lcl_showtopbar                    & "') AS showtopbar, "
     sSQL = sSQL & " isnull(topbar_bgcolor,'"                & lcl_topbar_bgcolor                & "') AS topbar_bgcolor, "
     sSQL = sSQL & " isnull(topbar_fonttype,'"               & lcl_topbar_fonttype               & "') AS topbar_fonttype, "
     sSQL = sSQL & " isnull(topbar_fontcolor,'"              & lcl_topbar_fontcolor              & "') AS topbar_fontcolor, "
     sSQL = sSQL & " isnull(topbar_fontcolorhover,'"         & lcl_topbar_fontcolorhover         & "') AS topbar_fontcolorhover, "
     sSQL = sSQL & " isnull(showsidemenubar,'"               & lcl_showsidemenubar               & "') AS showsidemenubar, "
     sSQL = sSQL & " isnull(sidemenubar_alignment,'"         & lcl_sidemenubar_alignment         & "') AS sidemenubar_alignment, "
     sSQL = sSQL & " isnull(sidemenuoption_bgcolor,'"        & lcl_sidemenuoption_bgcolor        & "') AS sidemenuoption_bgcolor, "
     sSQL = sSQL & " isnull(sidemenuoption_bgcolorhover,'"   & lcl_sidemenuoption_bgcolorhover   & "') AS sidemenuoption_bgcolorhover, "
     sSQL = sSQL & " isnull(sidemenuoption_alignment,'"      & lcl_sidemenuoption_alignment      & "') AS sidemenuoption_alignment, "
     sSQL = sSQL & " isnull(sidemenuoption_fonttype,'"       & lcl_sidemenuoption_fonttype       & "') AS sidemenuoption_fonttype, "
     sSQL = sSQL & " isnull(sidemenuoption_fontcolor,'"      & lcl_sidemenuoption_fontcolor      & "') AS sidemenuoption_fontcolor, "
     sSQL = sSQL & " isnull(sidemenuoption_fontcolorhover,'" & lcl_sidemenuoption_fontcolorhover & "') AS sidemenuoption_fontcolorhover, "
     sSQL = sSQL & " isnull(showpageheader,'"                & lcl_showpagehader                 & "') AS showpageheader, "
     sSQL = sSQL & " isnull(pageheader_alignment,'"          & lcl_pageheader_alignment          & "') AS pageheader_alignment, "
     sSQL = sSQL & " isnull(pageheader_fontsize,'"           & lcl_pageheader_fontsize           & "') AS pageheader_fontsize, "
     sSQL = sSQL & " isnull(pageheader_fontcolor,'"          & lcl_pageheader_fontcolor          & "') AS pageheader_fontcolor, "
     sSQL = sSQL & " isnull(pageheader_fonttype,'"           & lcl_pageheader_fonttype           & "') AS pageheader_fonttype, "
     sSQL = sSQL & " isnull(pageheader_bgcolor,'"            & lcl_pageheader_bgcolor            & "') AS pageheader_bgcolor, "
     sSQL = sSQL & " isnull(showfooter,'"                    & lcl_showfooter                    & "') AS showfooter, "
     sSQL = sSQL & " isnull(footer_bgcolor,'"                & lcl_footer_bgcolor                & "') AS footer_bgcolor, "
     sSQL = sSQL & " isnull(footer_fonttype,'"               & lcl_footer_fonttype               & "') AS footer_fonttype, "
     sSQL = sSQL & " isnull(footer_fontcolor,'"              & lcl_footer_fontcolor              & "') AS footer_fontcolor, "
     sSQL = sSQL & " isnull(footer_fontcolorhover,'"         & lcl_footer_fontcolorhover         & "') AS footer_fontcolorhover, "
     sSQL = sSQL & " showRSS, "
     sSQL = sSQL & " isnull(url_twitter,'"                   & lcl_url_twitter                   & "') AS url_twitter, "
     sSQL = sSQL & " isnull(url_facebook,'"                  & lcl_url_facebook                  & "') AS url_facebook, "
     sSQL = sSQL & " isnull(url_myspace,'"                   & lcl_url_myspace                   & "') AS url_myspace, "
     sSQL = sSQL & " isnull(url_blogger,'"                   & lcl_url_blogger                   & "') AS url_blogger "
     sSQL = sSQL & " FROM egov_communitylink "
     sSQL = sSQL & " WHERE orgid = " & p_orgid

     set oCLInfo = Server.CreateObject("ADODB.Recordset")
    	oCLInfo.Open sSQL, Application("DSN"), 3, 1

     if not oCLInfo.eof then
        lcl_isEgovHomePage                = oCLInfo("isEgovHomePage")
        lcl_website_size                  = oCLInfo("website_size")
        lcl_website_size_customsize       = oCLInfo("website_size_customsize")
        lcl_website_alignment             = oCLInfo("website_alignment")
        lcl_website_bgcolor               = oCLInfo("website_bgcolor")
        lcl_showlogo                      = oCLInfo("showlogo")
        lcl_logo_filename                 = oCLInfo("logo_filename")
        lcl_logo_filenamebg               = oCLInfo("logo_filenamebg")
        lcl_logo_alignment                = oCLInfo("logo_alignment")
        lcl_showtopbar                    = oCLInfo("showtopbar")
        lcl_topbar_bgcolor                = oCLInfo("topbar_bgcolor")
        lcl_topbar_fonttype               = oCLInfo("topbar_fonttype")
        lcl_topbar_fontcolor              = oCLInfo("topbar_fontcolor")
        lcl_topbar_fontcolorhover         = oCLInfo("topbar_fontcolorhover")
        lcl_showsidemenubar               = oCLInfo("showsidemenubar")
        lcl_sidemenubar_alignment         = oCLInfo("sidemenubar_alignment")
        lcl_sidemenuoption_bgcolor        = oCLInfo("sidemenuoption_bgcolor")
        lcl_sidemenuoption_bgcolorhover   = oCLInfo("sidemenuoption_bgcolorhover")
        lcl_sidemenuoption_alignment      = oCLInfo("sidemenuoption_alignment")
        lcl_sidemenuoption_fonttype       = oCLInfo("sidemenuoption_fonttype")
        lcl_sidemenuoption_fontcolor      = oCLInfo("sidemenuoption_fontcolor")
        lcl_sidemenuoption_fontcolorhover = oCLInfo("sidemenuoption_fontcolorhover")
        lcl_showpageheader                = oCLInfo("showpageheader")
        lcl_pageheader_alignment          = oCLInfo("pageheader_alignment")
        lcl_pageheader_fontsize           = oCLInfo("pageheader_fontsize")
        lcl_pageheader_fontcolor          = oCLInfo("pageheader_fontcolor")
        lcl_pageheader_fonttype           = oCLInfo("pageheader_fonttype")
        lcl_pageheader_bgcolor            = oCLInfo("pageheader_bgcolor")
        lcl_showfooter                    = oCLInfo("showfooter")
        lcl_footer_bgcolor                = oCLInfo("footer_bgcolor")
        lcl_footer_fonttype               = oCLInfo("footer_fonttype")
        lcl_footer_fontcolor              = oCLInfo("footer_fontcolor")
        lcl_footer_fontcolorhover         = oCLInfo("footer_fontcolorhover")
        lcl_showRSS                       = oCLInfo("showRSS")
        lcl_url_twitter                   = oCLInfo("url_twitter")
        lcl_url_facebook                  = oCLInfo("url_facebook")
        lcl_url_myspace                   = oCLInfo("url_myspace")
        lcl_url_blogger                   = oCLInfo("url_blogger")
     end if

     oCLInfo.close
     set oCLInfo = nothing

  'else
    'Defaults
     'lcl_isEgovHomePage                = 0
     'lcl_website_size                  = getCLOptionDefault("WEBSITE_SIZE")
     'lcl_website_size_customsize       = ""
     'lcl_website_alignment             = getCLOptionDefault("WEBSITE_ALIGN")
     'lcl_website_bgcolor               = "ffffff"
     'lcl_showlogo                      = 1
     'lcl_logo_filename                 = ""
     'lcl_logo_filenamebg               = ""
     'lcl_logo_alignment                = getCLOptionDefault("WEBSITE_LOGO_ALIGN")
     'lcl_showtopbar                    = 1
     'lcl_topbar_fonttype               = getCLOptionDefault("TOPBAR_FONTTYPE")
     'lcl_topbar_bgcolor                = "ffffff"
     'lcl_topbar_fontcolor              = "000000"
     'lcl_topbar_fontcolorhover         = "000000"
     'lcl_showsidemenubar               = 1
     'lcl_sidemenubar_alignment         = getCLOptionDefault("SIDEMENUBAR_ALIGN")
     'lcl_sidemenuoption_bgcolor        = "efefef"
     'lcl_sidemenuoption_bgcolorhover   = "c0c0c0"
     'lcl_sidemenuoption_alignment      = getCLOptionDefault("SIDEMENUOPT_TEXTALIGN")
     'lcl_sidemenuoption_fonttype       = getCLOptionDefault("SIDEMENUOPT_FONTTYPE")
     'lcl_sidemenuoption_fontcolor      = "000000"
     'lcl_sidemenuoption_fontcolorhover = "000000"
     'lcl_showpageheader                = 1
     'lcl_pageheader_alignment          = getCLOptionDefault("PAGEHEADER_ALIGN")
     'lcl_pageheader_fontsize           = "12"
     'lcl_pageheader_fontcolor          = "000000"
     'lcl_pageheader_fonttype           = getCLOptionDefault("PAGEHEADER_FONTTYPE")
     'lcl_pageheader_bgcolor            = "efefef"
     'lcl_showfooter                    = 1
     'lcl_footer_fonttype               = getCLOptionDefault("FOOTER_FONTTYPE")
     'lcl_footer_bgcolor                = "ffffff"
     'lcl_footer_fontcolor              = "000000"
     'lcl_footer_fontcolorhover         = "000000"
  end if

end sub

'------------------------------------------------------------------------------
function getWebsiteWidth(iWebsiteSize, iWebsiteSizeCustom)

  lcl_return = "800"

  if iWebsiteSize <> "" then
     iWebsiteSize = UCASE(iWebsiteSize)

     if iWebsiteSize = "S" then
        lcl_return = "600"
     elseif iWebsiteSize = "M" then
        lcl_return = "800"
     elseif iWebsiteSize = "L" then
        lcl_return = "1000"
     elseif iWebsiteSize = "C" then
        if iWebsiteSizeCustom <> "" then
           lcl_customsize = iWebsiteSizeCustom
        else
           lcl_customsize = "800"
        end if

        lcl_return = lcl_customsize
     end if
  end if

  getWebsiteWidth = lcl_return

end function

'------------------------------------------------------------------------------
function getDefaultLogo(iLogoType, p_orgid)
  lcl_return = ""

  if iLogoType = "" then
     iLogoType = "LEFT"
  end if

  sSQL = "SELECT orgTopGraphic" & iLogoType & "URL AS orgLogoURL "
  sSQL = sSQL & " FROM organizations "
  sSQL = sSQL & " WHERE orgid = " & p_orgid

  set oOrgLogo = Server.CreateObject("ADODB.Recordset")
  oOrgLogo.Open sSQL, Application("DSN"), 3, 1

  if not oOrgLogo.eof then
     lcl_return = oOrgLogo("orgLogoURL")
  end if

  oOrgLogo.close
  set oOrgLogo = nothing

  getDefaultLogo = lcl_return

end function

'------------------------------------------------------------------------------
sub ShowLoggedinLinks(iOrgID)  'Found in public-side "include_top_functions.asp"

	'Manage Account Link
  buildTopBarLink "MANAGE ACCOUNT", "#communitylink_preview"

	'View Standard EGov Payments Link
 	if lcl_orghasfeature_payments AND lcl_publiccanviewfeature_payments then
     buildTopBarLink "VIEW PAYMENTS", "#communitylink_preview"
  end if

	'View Submitted Action Line Requests Link
  if lcl_orghasfeature_action_line AND lcl_publiccanviewfeature_action_line then
     buildTopBarLink "VIEW REQUESTS", "#communitylink_preview"
  end if

	'View Shopping Cart (Purchases) Link
 	if lcl_orghasfeature_activities AND lcl_publiccanviewfeature_activities then
     buildTopBarLink "VIEW CART", "#communitylink_preview"
	 end if

 	if (lcl_orghasfeature_facilities  AND lcl_publiccanviewfeature_facilities) _
  OR (lcl_orghasfeature_activities  AND lcl_publiccanviewfeature_activities) _
  OR (lcl_orghasfeature_memberships AND lcl_publiccanviewfeature_memberships) _
  OR (lcl_orghasfeature_gifts       AND lcl_publiccanviewfeature_gifts) then
      buildTopBarLink "VIEW PURCHASES", "#communitylink_preview"
  end if

 'View Bids (Bid Postings) Link
  if lcl_orghasfeature_bid_postings AND lcl_publiccanviewfeature_bid_postings then
     buildTopBarLink "VIEW BIDS", "#communitylink_preview"
  end if

	'Logout Link
  buildTopBarLink "LOGOUT", "#communitylink_preview"

end sub

'------------------------------------------------------------------------------
function PublicCanViewFeature(iOrgId,sFeature)  'Found in public-side "include_top_functions.asp"
	Dim sSql, oRs
 lcl_return = False

	sSQL = "SELECT FO.publiccanview  "
	sSQL = sSQL & " FROM egov_organizations_to_features FO, egov_organization_features F "
	sSQL = sSQL & " WHERE FO.featureid = F.featureid  "
	sSQL = sSQL & " AND orgid = "      & iOrgId
	sSQL = sSQL & " AND F.feature = '" & sFeature & "' "

	set oCanView = Server.CreateObject("ADODB.Recordset")
	oCanView.Open sSQL, Application("DSN"), 3, 1

	if not oCanView.eof then
  		lcl_return = oCanView("publiccanview")
 end if
	
	oCanView.close
	set oCanView = nothing 

 PublicCanViewFeature = lcl_return

End Function 

'------------------------------------------------------------------------------
sub buildTopBarLink(iLabel, iURL)

 	response.write "<a href=""" & iURL & """ class=""topBarOption"">" & iLabel & "</a>" & vbcrlf

  if iLabel <> "LOGOUT" then
     response.write " | " & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
function getState(p_orgid)
  lcl_return = ""

  if p_orgid <> "" then
     sSQL = "SELECT isnull(orgstate,'') as orgstate "
     sSQL = sSQL & " FROM organizations "
     sSQL = sSQL & " WHERE orgid = " & p_orgid

   		set oState = Server.CreateObject("ADODB.Recordset")
   		oState.Open sSQL, Application("DSN"), 3, 1

     if not oState.eof then
        lcl_return = oState("orgstate")
     end if

   		oState.close
   		set oState = nothing
  end if

end function

'------------------------------------------------------------------------------
function getCLFeatAvailCount(p_orgid)

  lcl_return = 0

  if p_orgid <> "" then
     sSQL = "SELECT count(f.featureid) as total_count "
     sSQL = sSQL & " FROM egov_organization_features f, egov_organizations_to_features otf "
     sSQL = sSQL & " WHERE otf.featureid = f.featureid "
     sSQL = sSQL & " AND f.CommunityLinkOn = 1 "
     sSQL = sSQL & " AND otf.orgid = " & p_orgid

     set oCLFeatureCnt = Server.CreateObject("ADODB.Recordset")
     oCLFeatureCnt.Open sSQL, Application("DSN"), 3, 1

     if not oCLFeatureCnt.eof then
        lcl_return = oCLFeatureCnt("total_count")
     end if

     oCLFeatureCnt.close
     set oCLFeatureCnt = nothing

  end if

  getCLFeatAvailCount = lcl_return

end function

'------------------------------------------------------------------------------
function checkForCommunityLinkFeature(p_orgid, p_featureid)
  lcl_return = false

  if p_orgid <> "" AND p_featureid <> "" then
     sSQL = "SELECT cl.cl_displayid "
     sSQL = sSQL & " FROM egov_communitylink_displayorgfeatures cl "
     sSQL = sSQL & " WHERE cl.orgid = "   & p_orgid
     sSQL = sSQL & " AND cl.featureid = " & p_featureid

     set oCLExists = Server.CreateObject("ADODB.Recordset")
     oCLExists.Open sSQL, Application("DSN"), 3, 1

     if not oCLExists.eof then
        lcl_return = true
     end if

     oCLExists.close
     set oCLExists = nothing
  end if

  checkForCommunityLinkFeature = lcl_return

end function

'------------------------------------------------------------------------------
sub setupColorSelection(p_fieldid, p_value, p_numLines)
  lcl_lineSeparator  = "&nbsp;"
  lcl_lineSeparator2 = ""

  if p_numLines = 2 then
     lcl_lineSeparator  = "<br />" & vbcrlf
  elseif p_numLines = 3 then
     lcl_lineSeparator  = "<br />" & vbcrlf
     lcl_lineSeparator2 = "<br />" & vbcrlf
  end if

  response.write "<input type=""text"" name=""" & p_fieldid & """ id=""" & p_fieldid & """ value=""" & p_value & """ size=""7"" maxlength=""6"" onchange=""changePreviewColor('" & p_fieldid & "')"" />" & vbcrlf
  response.write lcl_lineSeparator
  response.write "<a href=""javascript:openWin('../colorpalette.asp','" & p_fieldid & "','','','')"" id=""" & p_fieldid & "_selectcolorlink"">[select color]</a>" & vbcrlf
  response.write lcl_lineSeparator2
  'response.write "<div id=""" & p_fieldid & "_previewcolor"" bgcolor=""" & p_value & """ style=""display:inline; border:1px solid #000000;"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div>" & vbcrlf
  response.write "<div id=""" & p_fieldid & "_previewcolor"" style=""background-color:#" & p_value & ";display:inline; border:1px solid #000000;"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub getSocialSiteIcons(p_orientation, p_showRSS, p_twitter, p_facebook, p_myspace, p_blogger)

 'Determine the layout orientation for the icons
 'H = horizontal
 'V = vertical
  if p_orientation <> "" then
     lcl_orientation = UCASE(p_orientation)
  else
     lcl_orientation = "H"
  end if

  if lcl_orientation = "V" then
     'lcl_icon_separation = "<br />" & vbcrlf
     lcl_icon_separation = "</tr><tr>" & vbcrlf
  else
     lcl_icon_separation = ""
  end if

  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""4"" style=""font-size:9px;"">" & vbcrlf
  response.write "  <caption>Follow us on:</caption>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf

 'Twitter
  if p_twitter <> "" then
     'response.write "      <td id=""icon_twitter"" align=""center"" nowrap=""nowrap"" onmouseover=""document.getElementById('icon_twitter').style.border='1pt solid #000000';"" onmouseout=""document.getElementById('icon_twitter').style.border='0pt solid #000000';"">" & vbcrlfresponse.write "      <td id=""icon_twitter"" align=""center"" nowrap=""nowrap"" onmouseover=""document.getElementById('icon_twitter').style.border='1pt solid #000000';"" onmouseout=""document.getElementById('icon_twitter').style.border='0pt solid #000000';"">" & vbcrlf
     response.write "      <td id=""icon_twitter"" align=""center"" nowrap=""nowrap"">" & vbcrlf
     response.write "          <a href=""" & p_twitter & """ target=""_twitter"">" & vbcrlf
     response.write "          <img src=""images/socialsites/icon_twitter.png"" border=""0"" alt=""Follow us on Twitter"" />" & vbcrlf
     response.write "          </a><br />" & vbcrlf
     response.write "Twitter" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write lcl_icon_separation
  end if

 'Facebook
  if p_facebook <> "" then
     response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
     response.write "          <a href=""" & p_facebook & """ target=""_facebook"">" & vbcrlf
     response.write "          <img src=""images/socialsites/icon_facebook.png"" border=""0"" alt=""Follow us on Facebook"" />" & vbcrlf
     response.write "          </a><br />" & vbcrlf
     response.write "Facebook" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write lcl_icon_separation
  end if

 'MySpace
  if p_myspace <> "" then
     response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
     response.write "          <a href=""" & p_myspace & """ target=""_myspace"">" & vbcrlf
     response.write "          <img src=""images/socialsites/icon_myspace.png"" border=""0"" alt=""Follow us on MySpace"" />" & vbcrlf
     response.write "          </a><br />" & vbcrlf
     response.write "MySpace" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write lcl_icon_separation
  end if

 'Blogger
  if p_blogger <> "" then
     response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
     response.write "          <a href=""" & p_blogger & """ target=""_blogger"">" & vbcrlf
     response.write "          <img src=""images/socialsites/icon_blogger.png"" border=""0"" alt=""Follow us on Blogger"" />" & vbcrlf
     response.write "          </a><br />" & vbcrlf
     response.write "Blogger" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write lcl_icon_separation
  end if

 'Show RSS
  if p_showRSS then
     response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
     response.write "          <a href=""" & session("egovclientwebsiteurl") & "/rssfeeds.asp"">" & vbcrlf
     response.write "          <img src=""images/socialsites/icon_rss.png"" border=""0"" alt=""Subscribe to our RSS Feeds"" />" & vbcrlf
     response.write "          </a><br />" & vbcrlf
     response.write "RSS" & vbcrlf
     response.write "      </td>" & vbcrlf
  end if

  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub saveCommunityLinkOption(ByVal p_orgid, ByVal p_feature, ByVal p_columnname, ByVal p_value, ByVal p_isAjaxRoutine, ByRef lcl_success)
  lcl_return    = ""
  lcl_value     = "NULL"
  lcl_featureid = 0
  lcl_success   = "N"

  if p_value = "" OR isnull(p_value) then
     lcl_value = "NULL"
  else
     lcl_value = "'" & dbsafe(p_value) & "'"
  end if

  if p_feature <> "" then
     lcl_featureid = getFeatureID(p_feature)
  end if

  if p_feature <> "" AND p_columnname <> "" then
     sSQL = "UPDATE egov_organizations_to_features SET "
     sSQL = sSQL & p_columnname & " = " & lcl_value
     sSQL = sSQL & " WHERE orgid = "   & p_orgid
     sSQL = sSQL & " AND featureid = " & lcl_featureid

    	set oUpdateCLOption = Server.CreateObject("ADODB.Recordset")
   	 oUpdateCLOption.Open sSQL, Application("DSN"), 3, 1

     set oUpdateCLOption = nothing

     lcl_success = "Y"
  end if

end sub

'------------------------------------------------------------------------------
sub displayViewRowLink(p_label, p_fieldID, p_URL, p_onMouseOver, p_onMouseOut, p_openNewWin)
  lcl_target      = " target=""_blank"""
  lcl_onMouseOver = ""
  lcl_onMouseOut  = ""

  if p_openNewWin <> "" then
     lcl_openNewWin = UCASE(p_openNewWin)
  else
     lcl_openNewWin = "Y"
  end if

  if lcl_openNewWin <> "Y" then
     lcl_target = ""
  end if

  if p_onMouseOver <> "" then
     lcl_onMouseOver = " onmouseover=""" & p_onmouseover & """"
  end if

  if p_onMouseOut <> "" then
     lcl_onMouseOut = " onmouseout=""" & p_onmouseout & """"
  end if

  response.write "<a href=""" & p_URL & """" & lcl_target & ">" & vbcrlf
  response.write "<span name=""" & p_fieldID & """ id=""" & p_fieldID & """" & lcl_onMouseOver & lcl_onMouseOut & " style=""cursor:pointer"">" & vbcrlf
  response.write   p_label & vbcrlf
  response.write "</span>" & vbcrlf
  response.write "</a>" & vbcrlf

end sub

'------------------------------------------------------------------------------
function getFeaturePortalType(p_featureid)
  lcl_return = ""

  if p_featureid <> "" then
     sSQL = "SELECT CL_portaltype "
     sSQL = sSQL & " FROM egov_organization_features "
     sSQL = sSQL & " WHERE featureid = " & p_featureid

     set oGetPortalType = Server.CreateObject("ADODB.Recordset")
     oGetPortalType.Open sSQL, Application("DSN"), 3, 1

     if not oGetPortalType.eof then
        lcl_return = trim(oGetPortalType("CL_portalType"))
     end if

     oGetPortalType.close
     set oGetPortalType = nothing
  end if

  getFeaturePortalType = lcl_return

end function

'------------------------------------------------------------------------------
sub runInlineJavascripts(p_scripts)

  if p_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write p_scripts & vbcrlf
     response.write "</script>" & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
sub displayPortalSections(p_portalLayoutType, p_column_num, p_orgid, p_orgRegistration, p_userid, p_wrap_td_tags, p_column_width, p_showRSS)

  if p_column_num <> "" then
     lcl_columnNum = p_column_num
  else
     lcl_columnNum = "1"
  end if

  if p_wrap_td_tags <> "" then
     lcl_wrapTDTags = UCASE(p_wrap_td_tags)
  else
     lcl_wrapTDTags = "Y"
  end if

  if p_column_width <> "" then
     lcl_ColumnWidth = p_column_width
  else
     lcl_ColumnWidth = "100%"
  end if

  if p_portalLayoutType <> "" then
     lcl_portalLayoutType = UCASE(p_portalLayoutType)
  else
     lcl_portalLayoutType = "CL"
  end if

  if lcl_portalLayoutType = "SAVVY" then
     lcl_columncheck = "d.isSavvyOn"
  else
     lcl_columncheck = "d.isCommunityLinkOn"
  end if

  if p_showRSS <> "" then
     lcl_showRSS = p_showRSS
  else
     lcl_showRSS = "Y"
  end if

 'Retrieve all of the features for the column specified
  sSQL = " SELECT d.orgid, d.featureid, d.featurename, d.portalcolumn, d.displayorder, "
  sSQL = sSQL & lcl_columncheck & ", "
  sSQL = sSQL & " isnull(d.numListItemsShown_"           & lcl_portalLayoutType & ",1) AS numListItemsShown, "
  sSQL = sSQL & " isnull(d.sectionheader_bgcolor_"       & lcl_portalLayoutType & ",'ffffff') AS sectionheader_bgcolor, "
  sSQL = sSQL & " isnull(d.sectionheader_linecolor_"     & lcl_portalLayoutType & ",'000000') AS sectionheader_linecolor, "
  sSQL = sSQL & " isnull(d.sectionheader_fonttype_"      & lcl_portalLayoutType & ",'" & getCLOptionDefault("SECTIONHEADER_FONTTYPE") & "') AS sectionheader_fonttype, "
  sSQL = sSQL & " isnull(d.sectionheader_fontcolor_"     & lcl_portalLayoutType & ",'000000') AS sectionheader_fontcolor, "
  sSQL = sSQL & " isnull(d.sectiontext_bgcolor_"         & lcl_portalLayoutType & ",'ffffff') AS sectiontext_bgcolor, "
  sSQL = sSQL & " isnull(d.sectiontext_fonttype_"        & lcl_portalLayoutType & ",'" & getCLOptionDefault("SECTIONTEXT_FONTTYPE") & "') AS sectiontext_fonttype, "
  sSQL = sSQL & " isnull(d.sectiontext_fontcolor_"       & lcl_portalLayoutType & ",'000000') AS sectiontext_fontcolor, "
  sSQL = sSQL & " isnull(d.sectionlinks_alignment_"      & lcl_portalLayoutType & ",'" & getCLOptionDefault("SECTIONLINKS_ALIGN")    & "') AS sectionlinks_alignment, "
  sSQL = sSQL & " isnull(d.sectionlinks_fonttype_"       & lcl_portalLayoutType & ",'" & getCLOptionDefault("SECTIONLINKS_FONTTYPE") & "') AS sectionlinks_fonttype, "
  sSQL = sSQL & " isnull(d.sectionlinks_fontcolor_"      & lcl_portalLayoutType & ",'000000') AS sectionlinks_fontcolor, "
  sSQL = sSQL & " isnull(d.sectionlinks_fontcolorhover_" & lcl_portalLayoutType & ",'000000') AS sectionlinks_fontcolorhover "
  sSQL = sSQL & " FROM egov_communitylink_displayorgfeatures d, "
  sSQL = sSQL &      " egov_organizations_to_features FO, "
  sSQL = sSQL &      " egov_organization_features f "
  sSQL = sSQL & " WHERE d.featureid = f.featureid "
  sSQL = sSQL & " AND F.featureid = FO.featureid "
  sSQL = sSQL & " AND FO.orgid = " & p_orgid
  sSQL = sSQL & " AND d.portalcolumn = " & lcl_ColumnNum
  sSQL = sSQL & " AND " & lcl_columncheck & " = 1 "
  sSQL = sSQL & " ORDER BY isnull(d.displayorder, 1), d.featurename "

  set oPortalColumns = Server.CreateObject("ADODB.Recordset")
  oPortalColumns.Open sSQL, Application("DSN"), 3, 1

  if not oPortalColumns.eof then

     if lcl_wrapTDTags = "Y" then
        if lcl_columnNum = 2 then
           lcl_td_bgcolor  = " background-color:#" & oPortalColumns("sectiontext_bgcolor")
           lcl_ColumnWidth = lcl_ColumnWidth+5
        else
           lcl_td_bgcolor = ""
        end if

        response.write "<td height=""100%"" style=""width:" & lcl_ColumnWidth & "px;" & lcl_td_bgcolor & """>" & vbcrlf
     end if

     i = 0
     do while not oPortalColumns.eof
        i = i + 1

       'Get the CL_portaltype
        lcl_portaltype = getFeaturePortalType(oPortalColumns("featureid"))

       'Build the Section Header styles
        lcl_sectionheader_styles = ""
        lcl_sectionheader_styles = lcl_sectionheader_styles & "font-size:11px;"
        lcl_sectionheader_styles = lcl_sectionheader_styles & "font-weight:bold;"
        lcl_sectionheader_styles = lcl_sectionheader_styles & "font-family:"              & oPortalColumns("sectionheader_fonttype")  & ";"
        lcl_sectionheader_styles = lcl_sectionheader_styles & "color:#"                   & oPortalColumns("sectionheader_fontcolor") & ";"
        lcl_sectionheader_styles = lcl_sectionheader_styles & "background-color:#"        & oPortalColumns("sectionheader_bgcolor")   & ";"
        lcl_sectionheader_styles = lcl_sectionheader_styles & "border-bottom:1pt solid #" & oPortalColumns("sectionheader_linecolor") & ";"
        lcl_sectionheader_styles = lcl_sectionheader_styles & "padding:5px;"

       'Show a top border if this is NOT the 1st section in the column.
        if i > 1 then
           lcl_sectionheader_styles = lcl_sectionheader_styles & "border-top:1pt solid #" & oPortalColumns("sectionheader_linecolor") & ";"
        end if

        if lcl_columnNum = 1 then
           'lcl_sectionheader_styles = lcl_sectionheader_styles & "margin-top:10px;"
           lcl_sectionheader_styles = lcl_sectionheader_styles & "padding-bottom:4px;"
        end if

       'Build the Section Text styles
        lcl_sectiontext_styles = ""
        lcl_sectiontext_styles = lcl_sectiontext_styles & "font-size:11px;"
        lcl_sectiontext_styles = lcl_sectiontext_styles & "font-family:"       & oPortalColumns("sectiontext_fonttype")  & ";"
        lcl_sectiontext_styles = lcl_sectiontext_styles & "color:#"            & oPortalColumns("sectiontext_fontcolor") & ";"
        lcl_sectiontext_styles = lcl_sectiontext_styles & "background-color:#" & oPortalColumns("sectiontext_bgcolor")   & ";"

        response.write "    <div>" & vbcrlf
        'response.write "      <div align=""left"" style=""" & lcl_sectionheader_styles & """>" & oPortalColumns("featurename") & "</div>" & vbcrlf
        response.write "      <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" style=""" & lcl_sectionheader_styles & """>" & vbcrlf
        response.write "        <tr>" & vbcrlf
        response.write "            <td align=""left"">" & oPortalColumns("featurename") & "</td>" & vbcrlf
        response.write "            <td align=""right"">" & vbcrlf
                                     if lcl_showRSS = "Y" then
                                        checkForRSSFeed p_orgid, oPortalColumns("featureid"),"", ""
                                     end if
        response.write "            </td>" & vbcrlf
        response.write "        </tr>" & vbcrlf
        response.write "      </table>" & vbcrlf
        response.write "      <div align=""left"" style=""" & lcl_sectiontext_styles   & """>" & vbcrlf
                                getPortalInfo lcl_portalLayoutType, p_orgid, oPortalColumns("featureid"), p_orgRegistration, p_userid, _
                                              oPortalColumns("numListItemsShown"), lcl_portaltype, _
                                              oPortalColumns("sectionheader_bgcolor"), oPortalColumns("sectiontext_bgcolor"), _
                                              oPortalColumns("sectiontext_fontcolor"), oPortalColumns("sectionlinks_alignment"), _
                                              oPortalColumns("sectionlinks_fonttype"), oPortalColumns("sectionlinks_fontcolor"), _
                                              oPortalColumns("sectionlinks_fontcolorhover")

        response.write "      </div>" & vbcrlf
        response.write "    </div>" & vbcrlf

        oPortalColumns.movenext
     loop

     if lcl_wrapTDTags = "Y" then
        response.write "</td>" & vbcrlf
     end if

  end if

  oPortalColumns.close
  set oPortalColumns = nothing  

end sub

'------------------------------------------------------------------------------
sub getPortalInfo(p_portalLayoutType, p_orgid, p_featureid, p_orgRegistration, p_userid, p_numListItemsShown, p_portaltype, _
                  p_sectionheader_bgcolor, p_sectiontext_bgcolor, p_sectiontext_fontcolor, p_sectionlinks_alignment, _
                  p_sectionlinks_fonttype, p_sectionlinks_fontcolor, p_sectionlinks_fontcolorhover)

  if p_portaltype <> "" then
     iPortalType = UCASE(p_portaltype)
  else
     iPortalType = ""
  end if

  if p_numListItemsShown = "" then
     iNumListItemsShown = 6
  else
     iNumListItemsShown = p_numListItemsShown
  end if

  if p_featureid <> "" then

    'Determine which portal section to show
    'The "portaltype" is set on the Org Feature Maintenance screen(s).
    'It is "CL_portaltype" on egov_organization_features.
     if iPortalType <> "" then
        select case UCASE(iPortalType)
          case "BLOG"
             displayBlogInfo p_orgid, p_featureid, p_orgRegistration, p_userid, iNumListItemsShown, p_sectionlinks_alignment, _
                             p_sectionlinks_fonttype, p_sectionlinks_fontcolor, p_sectionlinks_fontcolorhover

          case "COMMUNITY_CALENDAR"
             displayUpcomingEvents p_orgid, p_featureid, iNumListItemsShown, p_sectionheader_bgcolor, p_sectiontext_bgcolor, _
                                   p_sectiontext_fontcolor, p_sectionlinks_alignment, p_sectionlinks_fonttype, _
                                   p_sectionlinks_fontcolor, p_sectionlinks_fontcolorhover

          case "FAQ", "RUMORMILL"
             displayFAQ_RumorMill p_orgid, p_featureid, p_orgRegistration, p_userid, iNumListItemsShown, p_sectionheader_bgcolor, _
                                  p_sectiontext_bgcolor, p_sectiontext_fontcolor, p_sectionlinks_alignment, p_sectionlinks_fonttype, _
                                  p_sectionlinks_fontcolor, p_sectionlinks_fontcolorhover, UCASE(iPortalType), "N"

          case "NEWS"
             displayCurrentNews p_orgid, p_featureid, iNumListItemsShown, p_sectionheader_bgcolor, p_sectiontext_bgcolor, _
                                p_sectiontext_fontcolor, p_sectionlinks_alignment, p_sectionlinks_fonttype, _
                                p_sectionlinks_fontcolor, p_sectionlinks_fontcolorhover

          case "DOCUMENTS"
             displayNewDocuments p_orgid, p_featureid, iNumListItemsShown, p_sectionheader_bgcolor, p_sectiontext_bgcolor, _
                                 p_sectiontext_fontcolor, p_sectionlinks_alignment, p_sectionlinks_fonttype, _
                                 p_sectionlinks_fontcolor, p_sectionlinks_fontcolorhover

          case else
             response.write "&nbsp;" & vbcrlf
        end select
     else
        response.write "&nbsp;" & vbcrlf
     end if

  end if

end sub

'------------------------------------------------------------------------------
sub displayBlogInfo(p_orgid, p_featureid, p_orgRegistration, p_userid, p_numListItems, p_sectionlinks_alignment, p_sectionlinks_fonttype, _
                    p_sectionlinks_fontcolor, p_sectionlinks_fontcolorhover)

  iLineCnt              = 0
  iImgCount             = 0
  lcl_scripts_viewlinks = ""

  if p_numListItems <> "" then
     iNumListItems = p_numListItems
  else
     iNumListItems = 1
  end if

  if p_sectionlinks_alignment <> "" then
     iSectionLinks_Alignment = p_sectionlinks_alignment
  else
     iSectionLinks_Alignment = "RIGHT"
  end if

 'Set up the View Links onmouseover
  lcl_onmouseover_viewlinks = "'11',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'" & p_sectionlinks_fonttype       & "',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'" & p_sectionlinks_fontcolorhover & "',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'underline',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "''"

 'Set up the View Links onmouseout
  lcl_onmouseout_viewlinks = "'11',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'" & p_sectionlinks_fonttype  & "',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'" & p_sectionlinks_fontcolor & "',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'none',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "''"

 'Set up this portal section's specific styles
  lcl_section_spacing  = "padding:5px 5px 5px 10px;"
  lcl_styles_viewlinks = " style=""" & lcl_section_spacing & """"

 'Set up the View All and Post Comments urls
  lcl_viewAll_url         = session("egovclientwebsiteurl") & "/mayorsblog/mayorsblog.asp"
  lcl_postcomments_formid = getCommentsFormID(p_orgid, p_featureid, "")
  lcl_postcomments_url    = session("egovclientwebsiteurl") & "/action.asp?actionid=" & lcl_postcomments_formid

  sSQL = "SELECT TOP " & iNumListItems & " mb.blogid, mb.userid, mb.title, mb.article, mb.createdbyid, mb.createdbydate, "
  sSQL = sSQL & " isnull(u.imagefilename,'') AS imagefilename, u.firstname + ' ' + u.lastname as createdbyname "
  sSQL = sSQL & " FROM egov_mayorsblog mb, users u "
  sSQL = sSQL & " WHERE mb.userid = u.userid "
  sSQL = sSQL & " AND mb.isInactive = 0 "
  sSQL = sSQL & " AND mb.orgid = " & p_orgid
  sSQL = sSQL & " ORDER BY mb.createdbydate DESC "

  set oBlogInfo = Server.CreateObject("ADODB.Recordset")
  oBlogInfo.Open sSQL, Application("DSN"), 3, 1

  if not oBlogInfo.eof then
     do while not oBlogInfo.eof
        iLineCnt = iLineCnt + 1

        if len(trim(oBlogInfo("article"))) > 500 then
           lcl_article = left(trim(oBlogInfo("article")),500) & "..."
        else
           lcl_article = trim(oBlogInfo("article"))
        end if

        lcl_display_img = ""
        lcl_img_src     = session("egovclientwebsiteurl") & "/admin"
        lcl_img_border  = "/communitylink/images"
        lcl_img_scripts = ""
        lcl_tooltip     = ""

        if oBlogInfo("imagefilename") <> "" then
           iImgCount   = iImgCount + 1
           lcl_tooltip = oBlogInfo("createdbyname")

           lcl_display_img = lcl_display_img & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""float:left; margin-right:5px"">" & vbcrlf
           lcl_display_img = lcl_display_img & "  <tr><td colspan=""3""><img id=""blogimg_top_" & iImgCount & """ src=""" & lcl_img_src & lcl_img_border & "/blog_img_top.jpg"" height=""16"" alt=""" & lcl_tooltip & """ /></td></tr>" & vbcrlf
           lcl_display_img = lcl_display_img & "  <tr>" & vbcrlf
           lcl_display_img = lcl_display_img & "      <td><img id=""blogimg_left_"  & iImgCount & """ src=""" & lcl_img_src & lcl_img_border & "/blog_img_left.jpg"" width=""11"" alt=""" & lcl_tooltip & """ /></td>" & vbcrlf
           lcl_display_img = lcl_display_img & "      <td><img id=""blogimg_"       & iImgCount & """ name=""blogimg"" src=""" & lcl_img_src & "/custom/pub/" & session("virtualdirectory") & "/unpublished_documents" & oBlogInfo("imagefilename") & """ alt=""" & lcl_tooltip & """ /></td>" & vbcrlf
           lcl_display_img = lcl_display_img & "      <td><img id=""blogimg_right_" & iImgCount & """ src=""" & lcl_img_src & lcl_img_border & "/blog_img_right.jpg"" width=""15"" alt=""" & lcl_tooltip & """ /></td>" & vbcrlf
           lcl_display_img = lcl_display_img & "  </tr>" & vbcrlf
           lcl_display_img = lcl_display_img & "  <tr><td colspan=""3""><img id=""blogimg_bottom_" & iImgCount & """ src=""" & lcl_img_src & lcl_img_border & "/blog_img_bottom.jpg"" height=""16"" alt=""" & lcl_tooltip & """ /></td></tr>" & vbcrlf
           lcl_display_img = lcl_display_img & "</table>" & vbcrlf
        else
           iImgCount = iImgCount
        end if

       'Build the View Links URLs
        lcl_viewIndividual_url  = session("egovclientwebsiteurl") & "/mayorsblog/mayorsblog_info.asp?id=" & oBlogInfo("blogid")

       'If there are multiple rows then show the "dotted-line" separator
        if iLineCnt > 1 then
           lcl_section_separator = "border-top:1pt dotted #c0c0c0;"
        else
           lcl_section_separator = ""
        end if

        response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" style=""" & lcl_section_separator & lcl_section_spacing & """>" & vbcrlf
        response.write "  <tr>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write            lcl_display_img & vbcrlf
        response.write "          <p>" & vbcrlf
        response.write "             <strong style=""font-size:12px"">" & oBlogInfo("title") & "</strong><br />" & vbcrlf
        response.write "             <i style=""font-size:10px"">by: " & oBlogInfo("createdbyname") & " on " & FormatDateTime(oBlogInfo("createdbydate"),vbshortdate) & "</i>" & vbcrlf
        response.write "          </p>" & vbcrlf
        response.write "          <p>" & lcl_article & "</p>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "  <tr>" & vbcrlf
        response.write "      <td align=""" & iSectionLinks_Alignment & """ style=""padding:5px;"">" & vbcrlf

       'View Article
        displayViewRowLink "View Article", "viewIndividual_blog_" & oBlogInfo("blogid"), lcl_viewIndividual_url, _
                           "changeElementStyles('viewIndividual_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseover_viewlinks & ");", _
                           "changeElementStyles('viewIndividual_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseout_viewlinks  & ");", _
                           "Y"

        lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewIndividual_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

        response.write "&nbsp;|&nbsp;" & vbcrlf

       'View All
        displayViewRowLink "View All", "viewAll_blog_" & oBlogInfo("blogid"), lcl_viewAll_url, _
                           "changeElementStyles('viewAll_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseover_viewlinks & ");", _
                           "changeElementStyles('viewAll_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseout_viewlinks  & ");", _
                           "N"

        lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_blog_"        & oBlogInfo("blogid") & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

       'Post a Comment
       'Only allow the user to enter a comment if the user has logged in.
       ' - p_userid          = request.cookie("userid")
       ' - p_orgRegistration = sOrgRegistration (found in classOrganization)
        if p_orgRegistration AND p_userid <> "" AND p_userid <> "-1" AND lcl_postcomments_formid <> "" then
           response.write "&nbsp;|&nbsp;" & vbcrlf

          'Post a Comment
           displayViewRowLink "Post a Comment", "postComments_blog_" & oBlogInfo("blogid"), lcl_postcomments_url, _
                              "changeElementStyles('postComments_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseover_viewlinks & ");", _
                              "changeElementStyles('postComments_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseout_viewlinks  & ");", _
                              "N"

           lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('postComments_blog_"   & oBlogInfo("blogid") & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

        end if

        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "</table>" & vbcrlf

        oBlogInfo.movenext
     loop
  else
     response.write "<div align=""" & p_sectionlinks_alignment & """" & lcl_styles_viewlinks & ">" & vbcrlf

    'View All
     displayViewRowLink "View All", "viewAll_upcomingevents", lcl_viewAll_url, _
                        "changeElementStyles('viewAll_upcomingevents'," & lcl_onmouseover_viewlinks & ");", _
                        "changeElementStyles('viewAll_upcomingevents'," & lcl_onmouseout_viewlinks  & ");", _
                        "N"

     lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_upcomingevents'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

     response.write "</div>" & vbcrlf
  end if

  oBlogInfo.close
  set oBlogInfo = nothing

 'Check for any javascripts to run
  runInlineJavascripts lcl_scripts_viewlinks

end sub

'------------------------------------------------------------------------------
sub displayUpcomingEvents(p_orgid, p_featureid, p_numListItems, p_sectionheader_bgcolor, p_sectiontext_bgcolor, _
                          p_sectiontext_fontcolor, p_sectionlinks_alignment, p_sectionlinks_fonttype, p_sectionlinks_fontcolor, _
                          p_sectionlinks_fontcolorhover)

  iLineCnt              = 0
  lcl_scripts_viewlinks = ""

  if p_numListItems <> "" then
     iNumListItems = p_numListItems
  else
     iNumListItems = 5
  end if

  if p_sectionlinks_alignment <> "" then
     iSectionLinks_Alignment = p_sectionlinks_alignment
  else
     iSectionLinks_Alignment = "RIGHT"
  end if

 'Set up the View Links onmouseover
  lcl_onmouseover_viewlinks = "'11',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'" & p_sectionlinks_fonttype       & "',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'" & p_sectionlinks_fontcolorhover & "',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'underline',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "''"

 'Set up the View Links onmouseout
  lcl_onmouseout_viewlinks = "'11',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'" & p_sectionlinks_fonttype  & "',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'" & p_sectionlinks_fontcolor & "',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'none',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "''"

 'Set up this portal section's specific styles
  lcl_section_spacing         = "padding:5px 5px 5px 10px;"
  lcl_styles_container        = " style=""" & lcl_section_spacing & "cursor:pointer; border-bottom:1pt dotted #c0c0c0;"""
  lcl_styles_container_anchor = " style=""text-decoration:none; color:#" & p_sectiontext_bgcolor   & ";"""
  lcl_styles_eventdate        = " style=""font-size:10px; color:#"       & p_sectiontext_fontcolor & ";"""
  lcl_styles_subject          = " style=""font-size:11px; color:#"       & p_sectiontext_fontcolor & "; font-weight:bold;"""
  lcl_styles_viewlinks        = " style=""" & lcl_section_spacing & """"

 'Set up the View All url
  lcl_viewAll_url = session("egovclientwebsiteurl") & "/events/calendar.asp"

  sSQL = "SELECT TOP " & iNumListItems & " e.eventid, e.eventdate, e.subject "
  sSQL = sSQL & " FROM events e "
  sSQL = sSQL & " WHERE e.orgid = " & p_orgid
  sSQL = sSQL & " AND (calendarfeature = '' OR calendarfeature IS NULL) "
  sSQL = sSQL & " AND datediff(dd, '" & Date() & "',e.eventdate) >= 0 "
  sSQL = sSQL & " ORDER BY e.eventdate "

  set oUpcomingEvents = Server.CreateObject("ADODB.Recordset")
  oUpcomingEvents.Open sSQL, Application("DSN"), 3, 1

  if not oUpcomingEvents.eof then
     do while not oUpcomingEvents.eof
        iLineCnt = iLineCnt + 1

       'Set up the View Links onclick
        lcl_viewIndividual_url  = session("egovclientwebsiteurl") & "/events/calendarevents.asp?date=" & month(oUpcomingEvents("eventdate")) & "-" & day(oUpcomingEvents("eventdate")) & "-" & year(oUpcomingEvents("eventdate"))

        lcl_onmouseover_event = " onmouseover=""changeElementStyles('events_" & oUpcomingEvents("eventid") & "','','','" & p_sectiontext_fontcolor & "','underline','" & p_sectionheader_bgcolor & "');"""
        lcl_onmouseout_event  = " onmouseout=""changeElementStyles('events_"  & oUpcomingEvents("eventid") & "','','','','none','" & p_sectiontext_bgcolor & "');"""

        response.write "<a target=""_blank"" href=""" & lcl_viewIndividual_url & """" & lcl_styles_container_anchor & ">" & vbcrlf
        response.write "<div id=""events_" & oUpcomingEvents("eventid") & """" & lcl_styles_container & lcl_onmouseover_event & lcl_onmouseout_event & ">" & vbcrlf
        response.write "  <span" & lcl_styles_eventdate & ">" & oUpcomingEvents("eventdate") & "</span><br />" & vbcrlf
        response.write "  <span" & lcl_styles_subject   & ">" & oUpcomingEvents("subject")   & "</span>" & vbcrlf
        response.write "</div>" & vbcrlf
        response.write "</a>" & vbcrlf

        oUpcomingEvents.movenext
     loop
  end if

  oUpcomingEvents.close
  set oUpcomingEvents = nothing

  response.write "<div align=""" & p_sectionlinks_alignment & """" & lcl_styles_viewlinks & ">" & vbcrlf

 'View All
  displayViewRowLink "View All", "viewAll_upcomingevents", lcl_viewAll_url, _
                     "changeElementStyles('viewAll_upcomingevents'," & lcl_onmouseover_viewlinks & ");", _
                     "changeElementStyles('viewAll_upcomingevents'," & lcl_onmouseout_viewlinks  & ");", _
                     "N"

  lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_upcomingevents'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

  response.write "</div>" & vbcrlf

 'Check for any javascripts to run
  runInlineJavascripts lcl_scripts_viewlinks

end sub

'------------------------------------------------------------------------------
sub displayFAQ_RumorMill(p_orgid, p_featureid, p_orgRegistration, p_userid, iNumListItemsShown, p_sectionheader_bgcolor, _
                         p_sectiontext_bgcolor, p_sectiontext_fontcolor, p_sectionlinks_alignment, p_sectionlinks_fonttype, _
                         p_sectionlinks_fontcolor, p_sectionlinks_fontcolorhover, p_faqtype, p_showMouseOver)

  iLineCnt                 = 0
  lcl_scripts_viewlinks    = ""
  lcl_onmouseover_faqRumor = ""
  lcl_onmouseout_faqRumor  = ""

  if p_numListItems <> "" then
     iNumListItems = p_numListItems
  else
     iNumListItems = 5
  end if

  if p_sectionlinks_alignment <> "" then
     iSectionLinks_Alignment = p_sectionlinks_alignment
  else
     iSectionLinks_Alignment = "RIGHT"
  end if

  if p_faqtype <> "" then
     iFAQType = UCASE(p_faqtype)
  else
     iFAQType = "FAQ"  'FAQ
  end if

  if p_showMouseOver <> "" then
     iShowMouseOver = UCASE(p_showMouseOver)
  else
     iShowMouseOver = "Y"
  end if

 'Set up the View Links onmouseover
  lcl_onmouseover_viewlinks = "'11',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'" & p_sectionlinks_fonttype       & "',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'" & p_sectionlinks_fontcolorhover & "',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'underline',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "''"

 'Set up the View Links onmouseout
  lcl_onmouseout_viewlinks = "'11',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'" & p_sectionlinks_fonttype  & "',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'" & p_sectionlinks_fontcolor & "',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'none',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "''"

 'Set up this portal section's specific styles
  lcl_section_spacing  = "padding:5px 5px 5px 10px;"
  lcl_styles_container = " style=""" & lcl_section_spacing & "border-bottom:1pt dotted #c0c0c0;"""
  lcl_styles_faqRumor  = " style=""font-size:11px; color:#" & p_sectiontext_fontcolor & "; font-weight:bold;"""
  lcl_styles_viewlinks = " style=""" & lcl_section_spacing & """"

 'Set up the View All and Post Comments urls
  lcl_viewAll_url         = session("egovclientwebsiteurl") & "/faq.asp?faqtype=" & iFAQType
  lcl_postcomments_formid = getCommentsFormID(p_orgid, p_featureid, "")
  lcl_postcomments_url    = sEgovWebsiteURL & "/action.asp?actionid=" & lcl_postcomments_formid

  sSQL = "SELECT TOP " & iNumListItems & " FAQ.FaqID, FAQ.FaqQ, FAQ.faqA, isnull(faqcategoryname,'') AS faqcategoryname "
  sSQL = sSQL & " FROM FAQ "
 	sSQL = sSQL &      " LEFT OUTER JOIN faq_categories C ON C.faqcategoryid = faq.faqcategoryid "
  sSQL = sSQL &      " AND C.faqtype = faq.faqtype "
 	sSQL = sSQL & " WHERE faq.orgid = " & p_orgid
  sSQL = sSQL & " AND (internalonly = 0 OR internalonly is null) "
  'sSQL = sSQL & " AND datediff(dd, '" & Date() & "',isnull(publicationstart,'" & Date() & "')) >= 0 "
  sSQL = sSQL & " AND UPPER(faq.faqtype) = '" & iFAQType & "' "

  if iFAQType = "FAQ" then
     sSQL = sSQL & " OR (faq.faqtype = '' OR faq.faqtype IS NULL) "
  end if

 	sSQL = sSQL & " ORDER BY publicationstart DESC, displayorder, sequence"

  set oFAQRumorInfo = Server.CreateObject("ADODB.Recordset")
  oFAQRumorInfo.Open sSQL, Application("DSN"), 3, 1

  if not oFAQRumorInfo.eof then
     do while not oFAQRumorInfo.eof
        iLineCnt = iLineCnt + 1

       'Set up the View Links onclick
        'lcl_viewIndividual_url  = session("egovclientwebsiteurl") & "/faq.asp?faqtype=" & iFAQType

        if iShowMouseOver = "Y" then
           lcl_onmouseover_event = " onmouseover=""changeElementStyles('" & iFAQType & "_" & oFAQRumorInfo("faqid") & "','','','" & p_sectiontext_fontcolor & "','underline','" & p_sectionheader_bgcolor & "');"""
           lcl_onmouseout_event  = " onmouseout=""changeElementStyles('"  & iFAQType & "_" & oFAQRumorInfo("faqid") & "','','','','none','" & p_sectiontext_bgcolor & "');"""
        end if

        response.write "<div id=""" & iFAQType & "_" & oFAQRumorInfo("faqid") & """" & lcl_styles_container & lcl_onmouseover_faqRumor & lcl_onmouseout_faqRumor & ">" & vbcrlf
        response.write "  <li><span" & lcl_styles_faqRumor & ">" & oFAQRumorInfo("FaqQ") & "</span>" & vbcrlf
        response.write "</div>" & vbcrlf

        oFAQRumorInfo.movenext
     loop
  end if

  oFAQRumorInfo.close
  set oFAQRumorInfo = nothing

  response.write "<div align=""" & p_sectionlinks_alignment & """" & lcl_styles_viewlinks & ">" & vbcrlf

 'View All
  displayViewRowLink "View All", "viewAll_" & iFAQType, lcl_viewAll_url, _
                     "changeElementStyles('viewAll_" & iFAQType & "'," & lcl_onmouseover_viewlinks & ");", _
                     "changeElementStyles('viewAll_" & iFAQType & "'," & lcl_onmouseout_viewlinks  & ");", _
                     "N"

  lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_" & iFAQType & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

 'Submit a Rumor/Ask a Question (Post a Comment)
 'Only allow the user to enter a comment if the user has logged in.
 ' - p_userid          = request.cookie("userid")
 ' - p_orgRegistration = sOrgRegistration (found in classOrganization)
  if p_orgRegistration AND p_userid <> "" AND p_userid <> "-1" AND lcl_postcomments_formid <> "" then
     if iFAQType = "FAQ" then
        lcl_comments_label = "Ask a Question"
     else
        lcl_comments_label = "Submit a Rumor"
     end if

    'Post a Comment
    'Only allow the user to enter a comment if the user has logged in.
    ' - p_userid          = request.cookie("userid")
    ' - p_orgRegistration = sOrgRegistration (found in classOrganization)
     if p_orgRegistration AND p_userid <> "" AND p_userid <> "-1" then
        response.write "&nbsp;|&nbsp;" & vbcrlf

        displayViewRowLink lcl_comments_label, "postComments_" & iFAQType, lcl_postcomments_url, _
                           "changeElementStyles('postComments_" & iFAQType & "'," & lcl_onmouseover_viewlinks & ");", _
                           "changeElementStyles('postComments_" & iFAQType & "'," & lcl_onmouseout_viewlinks  & ");", _
                           "N"

        lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('postComments_" & iFAQType & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf
     end if
  end if

  response.write "</div>" & vbcrlf

 'Check for any javascripts to run
  runInlineJavascripts lcl_scripts_viewlinks

end sub

'------------------------------------------------------------------------------
sub displayCurrentNews(p_orgid, p_featureid, p_numListItems, p_sectionheader_bgcolor, p_sectiontext_bgcolor, _
                       p_sectiontext_fontcolor, p_sectionlinks_alignment, p_sectionlinks_fonttype, p_sectionlinks_fontcolor, _
                       p_sectionlinks_fontcolorhover)

  iLineCnt              = 0
  lcl_scripts_viewlinks = ""

  if p_numListItems <> "" then
     iNumListItems = p_numListItems
  else
     iNumListItems = 5
  end if

  if p_sectionlinks_alignment <> "" then
     iSectionLinks_Alignment = p_sectionlinks_alignment
  else
     iSectionLinks_Alignment = "RIGHT"
  end if

 'Set up the View Links onmouseover
  lcl_onmouseover_viewlinks = "'11',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'" & p_sectionlinks_fonttype       & "',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'" & p_sectionlinks_fontcolorhover & "',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'underline',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "''"

 'Set up the View Links onmouseout
  lcl_onmouseout_viewlinks = "'11',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'" & p_sectionlinks_fonttype  & "',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'" & p_sectionlinks_fontcolor & "',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'none',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "''"

 'Set up this portal section's specific styles
  lcl_section_spacing         = "padding:5px 5px 5px 10px;"
  lcl_styles_container        = " style=""" & lcl_section_spacing & "cursor:pointer; border-bottom:1pt dotted #c0c0c0;"""
  lcl_styles_container_anchor = " style=""text-decoration:none; color:#" & p_sectiontext_bgcolor   & ";"""
  lcl_styles_articledate      = " style=""font-size:10px; color:#"       & p_sectiontext_fontcolor & ";"""
  lcl_styles_itemtitle        = " style=""font-size:11px; color:#"       & p_sectiontext_fontcolor & "; font-weight:bold;"""
  lcl_styles_viewlinks        = " style=""" & lcl_section_spacing & """"

 'Set up the View All url
  lcl_viewAll_url = session("egovclientwebsiteurl") & "/news/news.asp"

  sSQL = "SELECT TOP " & iNumListItems & " newsitemid, itemtitle, isnull(publicationstart,itemdate) AS articledate "
  sSQL = sSQL & " FROM egov_news_items "
  sSQL = sSQL & " WHERE itemdisplay = 1 "
  sSQL = sSQL & " AND orgid = "  & p_orgid
  sSQL = sSQL & " AND datediff(dd, '" & date() & "',isnull(publicationstart,itemdate)) <= 0 "
  sSQL = sSQL & " AND (publicationend IS NULL OR datediff(dd, '" & date() & "',isnull(publicationend,'" & date() & "')) >= 0) "
  sSQL = sSQL & " ORDER BY isnull(publicationstart,itemdate) DESC "

  set oCurrentNews = Server.CreateObject("ADODB.Recordset")
  oCurrentNews.Open sSQL, Application("DSN"), 3, 1

  if not oCurrentNews.eof then
     do while not oCurrentNews.eof
        iLineCnt = iLineCnt + 1

       'Set up the View Links onclick
        lcl_viewIndividual_url  = session("egovclientwebsiteurl") & "/news/news_info.asp?id=" & oCurrentNews("newsitemid")

        lcl_onmouseover_event = " onmouseover=""changeElementStyles('news_" & oCurrentNews("newsitemid") & "','','','" & p_sectiontext_fontcolor & "','underline','" & p_sectionheader_bgcolor & "');"""
        lcl_onmouseout_event  = " onmouseout=""changeElementStyles('news_"  & oCurrentNews("newsitemid") & "','','','','none','" & p_sectiontext_bgcolor & "');"""

        response.write "<a href=""javascript:openWin('" & lcl_viewIndividual_url & "','','news',400,250);""" & lcl_styles_container_anchor & ">" & vbcrlf
        response.write "<div id=""news_" & oCurrentNews("newsitemid") & """" & lcl_styles_container & lcl_onmouseover_event & lcl_onmouseout_event & ">" & vbcrlf
        response.write "  <span" & lcl_styles_articledate & ">" & oCurrentNews("articledate") & "</span><br />" & vbcrlf
        response.write "  <span" & lcl_styles_itemtitle   & ">" & oCurrentNews("itemtitle")   & "</span>" & vbcrlf
        response.write "</div>" & vbcrlf
        response.write "</a>" & vbcrlf

        oCurrentNews.movenext
     loop

  end if

  oCurrentNews.close
  set oCurrentNews = nothing

  response.write "<div align=""" & p_sectionlinks_alignment & """" & lcl_styles_viewlinks & ">" & vbcrlf

 'View All
  displayViewRowLink "View All", "viewAll_currentNews", lcl_viewAll_url, _
                     "changeElementStyles('viewAll_currentNews'," & lcl_onmouseover_viewlinks & ");", _
                     "changeElementStyles('viewAll_currentNews'," & lcl_onmouseout_viewlinks  & ");", _
                     "N"

  lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_currentNews'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

  response.write "</div>" & vbcrlf

 'Check for any javascripts to run
  runInlineJavascripts lcl_scripts_viewlinks

end sub

'------------------------------------------------------------------------------
sub displayNewDocuments(p_orgid, p_featureid, iNumListItemsShown, p_sectionheader_bgcolor, p_sectiontext_bgcolor, _
                        p_sectiontext_fontcolor, p_sectionlinks_alignment, p_sectionlinks_fonttype, p_sectionlinks_fontcolor, _
                        p_sectionlinks_fontcolorhover)

  iLineCnt              = 0
  lcl_scripts_viewlinks = ""

  if p_numListItems <> "" then
     iNumListItems = p_numListItems
  else
     iNumListItems = 5
  end if

  if p_sectionlinks_alignment <> "" then
     iSectionLinks_Alignment = p_sectionlinks_alignment
  else
     iSectionLinks_Alignment = "RIGHT"
  end if

 'Set up the View Links onmouseover
  lcl_onmouseover_viewlinks = "'11',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'" & p_sectionlinks_fonttype       & "',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'" & p_sectionlinks_fontcolorhover & "',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "'underline',"
  lcl_onmouseover_viewlinks = lcl_onmouseover_viewlinks & "''"

 'Set up the View Links onmouseout
  lcl_onmouseout_viewlinks = "'11',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'" & p_sectionlinks_fonttype  & "',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'" & p_sectionlinks_fontcolor & "',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "'none',"
  lcl_onmouseout_viewlinks = lcl_onmouseout_viewlinks & "''"

 'Set up this portal section's specific styles
  lcl_section_spacing         = "padding:5px 5px 5px 10px;"
  lcl_styles_container        = " style=""" & lcl_section_spacing & "border-bottom:1pt dotted #c0c0c0;"""
  'lcl_styles_container_anchor = " style=""text-decoration:none; color:#" & p_sectiontext_fontcolor & ";"""
  lcl_styles_container_anchor = " style=""color:#" & p_sectiontext_fontcolor & ";"""
  lcl_styles_newDoc           = " style=""font-size:11px; color:#" & p_sectiontext_fontcolor & "; font-weight:bold;"""
  lcl_styles_viewlinks        = " style=""" & lcl_section_spacing & """"

 'Set up the View All url
  lcl_viewAll_url = session("egovclientwebsiteurl") & "/docs/menu/home.asp"

		sSQL = "SELECT TOP " & iNumListItems & " documentid, documenturl, documenttitle, dateadded "
  sSQL = sSQL & " FROM documents "
  sSQL = sSQL & " WHERE orgid = " & p_orgid
  sSQL = sSQL & " AND UPPER(documenturl) LIKE ('%/PUBLISHED_DOCUMENTS%') "
  sSQL = sSQL & " ORDER BY dateadded desc "

  set oNewDocs = Server.CreateObject("ADODB.Recordset")
  oNewDocs.Open sSQL, Application("DSN"), 3, 1

  if not oNewDocs.eof then
     do while not oNewDocs.eof
        iLineCnt = iLineCnt + 1

       'Set up the View Links onclick
        lcl_viewIndividual_url  = session("egovclientwebsiteurl") & "/admin"
        lcl_viewIndividual_url = lcl_viewIndividual_url & replace(oNewDocs("documenturl"),"/public_documents300","")

        lcl_onmouseover_event = " onmouseover=""changeElementStyles('newdoc_" & oNewDocs("documentid") & "','','','" & p_sectiontext_fontcolor & "','underline','" & p_sectionheader_bgcolor & "');"""
        lcl_onmouseout_event  = " onmouseout=""changeElementStyles('newdoc_"  & oNewDocs("documentid") & "','','','','none','" & p_sectiontext_bgcolor & "');"""

        response.write "<a target=""_blank"" href=""" & lcl_viewIndividual_url & """" & lcl_styles_container_anchor & ">" & vbcrlf
        response.write "<div id=""newdoc_" & oNewDocs("documentid") & """" & lcl_styles_container & lcl_onmouseover_newDoc & lcl_onmouseout_newDoc & ">" & vbcrlf
        response.write "  <li><span" & lcl_styles_newDoc & ">" & oNewDocs("documenttitle") & "</span>" & vbcrlf
        response.write "</div>" & vbcrlf
        response.write "</a>" & vbcrlf

        oNewDocs.movenext
     loop
  end if

  oNewDocs.close
  set oNewDocs = nothing

  response.write "<div align=""" & p_sectionlinks_alignment & """" & lcl_styles_viewlinks & ">" & vbcrlf

 'View All
  displayViewRowLink "View All", "viewAll_newDoc", lcl_viewAll_url, _
                     "changeElementStyles('viewAll_newdoc'," & lcl_onmouseover_viewlinks & ");", _
                     "changeElementStyles('viewAll_newdoc'," & lcl_onmouseout_viewlinks  & ");", _
                     "N"

  lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_newdoc'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

  response.write "</div>" & vbcrlf

 'Check for any javascripts to run
  runInlineJavascripts lcl_scripts_viewlinks

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1

end sub
%>