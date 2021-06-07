<%
'Check for org features
 lcl_orghasfeature_registration       = orghasfeature(iorgid,"registration")
 lcl_orghasfeature_administrationlink = orghasfeature(iorgid,"AdministrationLink")
 lcl_orghasfeature_payments           = orghasfeature(iorgid,"payments")
 lcl_orghasfeature_action_line        = orghasfeature(iorgid,"action line")
 lcl_orghasfeature_activities         = orghasfeature(iorgid,"activities")
 lcl_orghasfeature_facilities         = orghasfeature(iorgid,"facilities")
 lcl_orghasfeature_memberships        = orghasfeature(iorgid,"memberships")
 lcl_orghasfeature_gifts              = orghasfeature(iorgid,"gifts")
 lcl_orghasfeature_bid_postings       = orghasfeature(iorgid,"bid_postings")

'Check for PublicCanViewFeature
 lcl_publiccanviewfeature_payments     = publiccanviewfeature(iorgid,"payments")
 lcl_publiccanviewfeature_action_line  = publiccanviewfeature(iorgid,"action line")
 lcl_publiccanviewfeature_activities   = publiccanviewfeature(iorgid,"activities")
 lcl_publiccanviewfeature_facilities   = publiccanviewfeature(iorgid,"facilities")
 lcl_publiccanviewfeature_memberships  = publiccanviewfeature(iorgid,"memberships")
 lcl_publiccanviewfeature_gifts        = publiccanviewfeature(iorgid,"gifts")
 lcl_publiccanviewfeature_bid_postings = publiccanviewfeature(iorgid,"big postings")

'------------------------------------------------------------------------------
function getImgBaseURL(p_URL)

  lcl_return = p_URL

  if request.servervariables("HTTPS") = "on" then
	    lcl_return = replace(p_URL,"http://www.egovlink.com","https://secure.egovlink.com")
  end if

  getImgBaseURL = lcl_return

end function

'------------------------------------------------------------------------------
sub displaySideMenubar(p_orgid, _
                       iSideMenuOptionBGColor, _
                       iSideMenuOptionBGColorHover, _
                       iSideMenuOptionAlignment, _
                       p_userid, _
                       p_isEgovHomePage)

  if p_isEgovHomePage <> "" then
     lcl_isEgovHomePage = p_isEgovHomePage
  else
     lcl_isEgovHomePage = 0
  end if

 'City Home (maintained in Org Properites)
  if oOrg.checkMenuOptionEnabled("CITY") then
     'response.write "<li><a href=""" & oOrg.GetOrgURL() & """>" & lcl_label_city & "</a></li>" & vbcrlf
     lcl_label_city = oOrg.getMenuOptionLabel("CITY")

     displaySideMenubarOption "0", iSideMenuOptionAlignment, oOrg.GetOrgURL(),  lcl_label_city
  end if

 'E-Gov Home (maintained in Org Properties)
  if oOrg.checkMenuOptionEnabled("EGOV") then
     'response.write "<li><a href=""" & oOrg.GetEgovURL() & """>" & lcl_label_egov & "</a></li>" & vbcrlf
     lcl_label_egov = oOrg.getMenuOptionLabel("EGOV")

     displaySideMenubarOption "1", iSideMenuOptionAlignment, oOrg.GetEgovURL(), lcl_label_egov
  end if

  'displaySideMenubarOption "0", iSideMenuOptionAlignment, oOrg.GetOrgURL(),  oOrg.GetOrgDisplayName("homewebsitetag")
  'displaySideMenubarOption "1", iSideMenuOptionAlignment, oOrg.GetEgovURL(), "E-Gov Home"
  'displaySideMenubarOption "0", iSideMenuOptionAlignment, oOrg.GetOrgURL(),  lcl_label_city
  'displaySideMenubarOption "1", iSideMenuOptionAlignment, oOrg.GetEgovURL(), lcl_label_egov

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

  if not oSideNav.eof then
   		do while not oSideNav.eof
       'Only display the menu option if:
       '1. the option (feature) does NOT equal "communitylink".
       '2. the option (featuer) DOES equal "communitylink" AND it has NOT been set to be the "e-gov home" page.
        if UCASE(oSideNav("feature")) <> "COMMUNITYLINK" OR (UCASE(oSideNav("feature")) = "COMMUNITYLINK" and NOT lcl_isEgovHomePage) then
           i = i + 1

          'Build the menu option url
  	      		if ucase(left(oSideNav("publicURL"),4)) = "HTTP" then
         				'They have their own page to start from
          				sNav = oSideNav("publicURL")
           else
         				'Start from our page
    		      		sNav = oSideNav("OrgEgovWebsiteURL") & "/" & oSideNav("publicURL")
           end if

           displaySideMenubarOption i, iSideMenuOptionAlignment, sNav, oSideNav("featurename")
        end if

        oSideNav.movenext
     loop
  end if

		oSideNav.close
		set oSideNav = nothing

	'Add the login link for those that have this
  i = i + 1

		if lcl_orghasfeature_registration AND p_userid <> "" AND p_userid <> "-1" then
     displaySideMenubarOption i, iSideMenuOptionAlignment, "logout.asp", "Logout"
  else
     displaySideMenubarOption i, iSideMenuOptionAlignment, "user_login.asp", "Login"
  end if

end sub

'------------------------------------------------------------------------------
sub displaySideMenubarOption(pID, _
                             p_alignment, _
                             p_url, _
                             p_label)

  lcl_onmouseover = " onmouseover=""setupMenuOption('OVER','" & pID & "');"""
  lcl_onmouseout  = " onmouseout=""setupMenuOption('OUT','"   & pID & "');"""
  lcl_onclick     = " onclick=""window.top.location.href='" & p_URL & "';"""

  response.write "<div id=""sideMenuBar" & pID & """ class=""sideMenuBar"" align=""" & p_alignment & """" & lcl_onmouseover & lcl_onmouseout & lcl_onclick & ">" & vbcrlf
  response.write "<a href=""" & p_URL & """ id=""sideMenuBarOption" & pID & """ class=""sideMenuBarOption"" target=""_top"">" & p_label & "</a>" & vbcrlf
  response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub ShowPublicDefaultFooterNav(p_orgid, _
                               iCount, _
                               p_isEgovHomePage)
		Dim sSQL, oNav, sNav, iTotalCount

  if p_isEgovHomePage <> "" then
     lcl_isEgovHomePage = p_isEgovHomePage
  else
     lcl_isEgovHomePage = 0
  end if

		sSQL = "SELECT O.OrgEgovWebsiteURL, "
  sSQL = sSQL & " isnull(FO.publicurl,F.publicURL) as publicURL, "
		sSQL = sSQL & " isnull(FO.featurename,F.featurename) as featurename, "
  sSQL = sSQL & " f.feature "
		sSQL = sSQL & " FROM organizations O, "
  sSQL = sSQL &      " egov_organizations_to_features FO, "
  sSQL = sSQL &      " egov_organization_features F "
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

        if ucase(left(oFooter("publicURL"),4)) = "HTTP" Then 
      				'They have their own page to start from
    		   		sNav = oFooter("publicURL")
     			else
      				'start from our page
    		   		sNav = oFooter("OrgEgovWebsiteURL") & "/" & oFooter("publicURL")
        end if

     			response.write "<a href=""" & sNav & """ class=""footerOption"" target=""_top"">" & oFooter("featurename") & "</a>" & vbcrlf

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

  getOrgTagLine = lcl_return

end function

'------------------------------------------------------------------------------
sub displayCommunityLinkOptions(iOptionName, _
                                iCurrentValue)

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
        if iCurrentValue <> "" then
           if UCASE(iCurrentValue) = UCASE(oCLOptions("optionvalue")) then
              lcl_selected_cloption = " selected=""selected"""
           else
              lcl_selected_cloption = ""
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
function getCommunityLinkID(p_orgid, _
                            p_userid)
  dim oCLID
  lcl_return = 0

  if p_orgid <> "" then
     if isnumeric(p_orgid) then
       'First check to see if the org has a community link record.
        sSQL = "SELECT communitylinkid "
        sSQL = sSQL & " FROM egov_communitylink "
        sSQL = sSQL & " WHERE orgid = " & p_orgid

        set oCLID = Server.CreateObject("ADODB.Recordset")
       	oCLID.Open sSQL, Application("DSN"), 3, 1

        if not oCLID.eof then
           lcl_return = oCLID("communitylinkid")
        end if

        oCLID.close
        set oCLID = nothing
     end if
  end if

  if Clng(lcl_return) = Clng(0) then
     lcl_return = createCommunityLink(p_orgid,p_userid)
  end if

  getCommunityLinkID = lcl_return

end function

'------------------------------------------------------------------------------
function createCommunityLink(p_orgid, _
                             p_userid)
  lcl_return = 0

  if iOrgID <> "" AND p_userid <> "" then
    'First check to see if the org has a community link record.
     sSQL = "INSERT INTO egov_communitylink (orgid, lastmodifiedbyid, lastmodifiedbydate) VALUES ("
     sSQL = sSQL & p_orgid  & ", "
     sSQL = sSQL & p_userid & ", "
     sSQL = sSQL & "'" & dbsafe(ConvertDateTimetoTimeZone(p_orgid)) & "' "
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
sub getCommunityLinkInfo(ByVal iCommunityLinkID, _
                         ByVal p_orgid, _
                         ByRef lcl_isEgovHomePage, _
                         ByRef lcl_website_size, _
                         ByRef lcl_website_size_customsize, _
                         ByRef lcl_website_alignment, _
                         ByRef lcl_website_bgcolor, _
                         ByRef lcl_showlogo, _
                         ByRef lcl_logo_filename, _
                         ByRef lcl_logo_filenamebg, _
                         ByRef lcl_logo_alignment, _
                         ByRef lcl_showtopbar, _
                         ByRef lcl_topbar_bgcolor, _
                         ByRef lcl_topbar_fonttype, _
                         ByRef lcl_topbar_fontcolor, _
                         ByRef lcl_topbar_fontcolorhover, _
                         ByRef lcl_showsidemenubar, _
                         ByRef lcl_sidemenubar_alignment, _
                         ByRef lcl_sidemenuoption_bgcolor, _
                         ByRef lcl_sidemenuoption_bgcolorhover, _
                         ByRef lcl_sidemenuoption_alignment, _
                         ByRef lcl_sidemenuoption_fonttype, _
                         ByRef lcl_sidemenuoption_fontcolor, _
                         ByRef lcl_sidemenuoption_fontcolorhover, _
                         ByRef lcl_showpageheader, _
                         ByRef lcl_pageheader_alignment, _
                         ByRef lcl_pageheader_fontsize, _
                         ByRef lcl_pageheader_fontcolor, _
                         ByRef lcl_pageheader_fonttype, _
                         ByRef lcl_pageheader_bgcolor, _
                         ByRef lcl_showfooter, _
                         ByRef lcl_footer_bgcolor, _
                         ByRef lcl_footer_fonttype, _
                         ByRef lcl_footer_fontcolor, _
                         ByRef lcl_footer_fontcolorhover, _
                         ByRef lcl_showRSS, _
                         ByRef lcl_url_twitter, _
                         ByRef lcl_url_facebook, _
                         ByRef lcl_url_myspace, _
                         ByRef lcl_url_blogger )

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
  lcl_pageheader_fontsize           = "12"
  lcl_pageheader_alignment          = getCLOptionDefault("PAGEHEADER_ALIGN")
  lcl_pageheader_fontcolor          = "000000"
  lcl_pageheader_fonttype           = getCLOptionDefault("PAGEHEADER_FONTTYPE")
  lcl_pageheader_bgcolor            = "efefef"
  lcl_showfooter                    = 1
  lcl_footer_fonttype               = getCLOptionDefault("FOOTER_FONTTYPE")
  lcl_footer_bgcolor                = "ffffff"
  lcl_footer_fontcolor              = "000000"
  lcl_footer_fontcolorhover         = "000000"
  lcl_showRSS                       = 0
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
     sSQL = sSQL & " isnull(showpageheader,'"                & lcl_showpageheader                & "') AS showpageheader, "
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
     'lcl_pageheader_fontsize           = "11"
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
function getWebsiteWidth(iWebsiteSize, _
                         iWebsiteSizeCustom)

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
function getDefaultLogo(iLogoType, _
                        p_orgid)
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
sub ShowLoggedinLinks(sPath)  'Found in public-side "include_top_functions.asp"

	'Manage Account Link
  buildTopBarLink "MANAGE ACCOUNT", sPath & "manage_account.asp"

	'View Standard EGov Payments Link
 	if lcl_orghasfeature_payments AND lcl_publiccanviewfeature_payments then
     buildTopBarLink "VIEW PAYMENTS", sPath & "user_home.asp?trantype=1"
  end if

	'View Submitted Action Line Requests Link
  if lcl_orghasfeature_action_line AND lcl_publiccanviewfeature_action_line then
     buildTopBarLink "VIEW REQUESTS", sPath & "user_home.asp?trantype=0"
  end if

	'View Shopping Cart (Purchases) Link
 	if lcl_orghasfeature_activities AND lcl_publiccanviewfeature_activities then
     buildTopBarLink "VIEW CART", sPath & "classes/class_cart.asp"
	 end if

 	if (lcl_orghasfeature_facilities  AND lcl_publiccanviewfeature_facilities) _
  OR (lcl_orghasfeature_activities  AND lcl_publiccanviewfeature_activities) _
  OR (lcl_orghasfeature_memberships AND lcl_publiccanviewfeature_memberships) _
  OR (lcl_orghasfeature_gifts       AND lcl_publiccanviewfeature_gifts) then
      buildTopBarLink "VIEW PURCHASES", sPath & "purchases_report/purchases_list.asp"
  end if

 'View Bids (Bid Postings) Link
  if lcl_orghasfeature_bid_postings AND lcl_publiccanviewfeature_bid_postings then
     buildTopBarLink "VIEW BIDS", sPath & "view_bids.asp"
  end if

	'Logout Link
  buildTopBarLink "LOGOUT", sPath & "logout.asp"

end sub

'------------------------------------------------------------------------------
function PublicCanViewFeature(iOrgId, _
                              sFeature)  'Found in public-side "include_top_functions.asp"
	Dim sSql, oRs
 lcl_return = False

	sSQL = "SELECT FO.publiccanview  "
	sSQL = sSQL & " FROM egov_organizations_to_features FO, "
 sSQL = sSQL &      " egov_organization_features F "
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
sub buildTopBarLink(iLabel, _
                    iURL)

 	response.write "<a href=""" & iURL & """ class=""topBarOption"">" & iLabel & "</a>" & vbcrlf

  if iLabel <> "LOGOUT" AND iLabel <> "LOGIN" then
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
     sSQL = sSQL & " FROM egov_organization_features f, "
     sSQL = sSQL &      " egov_organizations_to_features otf "
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
function checkForCommunityLinkFeature(p_orgid, _
                                      p_featureid)
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
sub setupColorSelection(p_fieldid, _
                        p_value, _
                        p_numLines)
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
  response.write "<div id=""" & p_fieldid & "_previewcolor"" bgcolor=""" & p_value & """ style=""display:inline; border:1px solid #000000;"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub getSocialSiteIcons(p_orientation, _
                       p_showRSS, _
                       p_twitter, _
                       p_facebook, _
                       p_myspace, _
                       p_blogger)

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

  if trim(p_twitter) <> "" OR trim(p_facebook) <> "" OR trim(p_myspace) <> "" OR trim(p_blogger) <> "" then
     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""4"" style=""font-size:9px;"">" & vbcrlf
     response.write "  <caption>Follow us on:</caption>" & vbcrlf
     response.write "  <tr valign=""top"">" & vbcrlf

    'Twitter
     if p_twitter <> "" then
        'response.write "      <td id=""icon_twitter"" align=""center"" nowrap=""nowrap"" onmouseover=""document.getElementById('icon_twitter').style.border='1pt solid #000000';"" onmouseout=""document.getElementById('icon_twitter').style.border='0pt solid #000000';"">" & vbcrlfresponse.write "      <td id=""icon_twitter"" align=""center"" nowrap=""nowrap"" onmouseover=""document.getElementById('icon_twitter').style.border='1pt solid #000000';"" onmouseout=""document.getElementById('icon_twitter').style.border='0pt solid #000000';"">" & vbcrlf
        response.write "      <td id=""icon_twitter"" align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <a href=""" & p_twitter & """ target=""_twitter"">" & vbcrlf
        response.write "          <img src=""" & replace(sEgovWebsiteURL,"http:","https:") & "/images/communitylink/socialsites/icon_twitter.png"" border=""0"" alt=""Follow us on Twitter"" />" & vbcrlf
        response.write "          </a><br />" & vbcrlf
        response.write "Twitter" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write lcl_icon_separation
     end if

    'Facebook
     if p_facebook <> "" then
        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <a href=""" & p_facebook & """ target=""_facebook"">" & vbcrlf
        response.write "          <img src=""" & replace(sEgovWebsiteURL,"http:","https:") & "/images/communitylink/socialsites/icon_facebook.png"" border=""0"" alt=""Follow us on Facebook"" />" & vbcrlf
        response.write "          </a><br />" & vbcrlf
        response.write "Facebook" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write lcl_icon_separation
     end if

    'MySpace
     if p_myspace <> "" then
        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <a href=""" & p_myspace & """ target=""_myspace"">" & vbcrlf
        response.write "          <img src=""" & replace(sEgovWebsiteURL,"http:","https:") & "/images/communitylink/socialsites/icon_myspace.png"" border=""0"" alt=""Follow us on MySpace"" />" & vbcrlf
        response.write "          </a><br />" & vbcrlf
        response.write "MySpace" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write lcl_icon_separation
     end if

    'Blogger
     if p_blogger <> "" then
        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <a href=""" & p_blogger & """ target=""_blogger"">" & vbcrlf
        response.write "          <img src=""" & replace(sEgovWebsiteURL,"http:","https:") & "/images/communitylink/socialsites/icon_blogger.png"" border=""0"" alt=""Follow us on Blogger"" />" & vbcrlf
        response.write "          </a><br />" & vbcrlf
        response.write "Blogger" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write lcl_icon_separation
     end if

    'Show RSS
     if p_showRSS then
        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <a href=""" & sEgovWebsiteURL & "/rssfeeds.asp"" target=""_top"">" & vbcrlf
        response.write "          <img src=""" & replace(sEgovWebsiteURL,"http:","https:") & "/images/communitylink/socialsites/icon_rss.png"" border=""0"" alt=""Subscribe to our RSS Feeds"" />" & vbcrlf
        response.write "          </a><br />" & vbcrlf
        response.write "RSS" & vbcrlf
        response.write "      </td>" & vbcrlf
     end if

     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
sub saveCommunityLinkOption(ByVal p_orgid, _
                            ByVal p_feature, _
                            ByVal p_columnname, _
                            ByVal p_value, _
                            ByVal p_isAjaxRoutine, _
                            ByRef lcl_success)
  lcl_return    = ""
  lcl_value     = "NULL"
  lcl_featureid = 0

  if p_value = "" then
     lcl_value = "NULL"
  else
     lcl_value = p_value
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
sub displayViewRowLink(p_label, _
                       p_fieldID, _
                       p_URL, _
                       p_onMouseOver, _
                       p_onMouseOut, _
                       p_openNewWin)

  lcl_target      = " target=""_top"""
  lcl_onMouseOver = ""
  lcl_onMouseOut  = ""

  if p_openNewWin <> "" then
     lcl_openNewWin = UCASE(p_openNewWin)
  else
     lcl_openNewWin = "Y"
  end if

  if lcl_openNewWin = "Y" then
     lcl_target = " target=""_blank"""
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
sub displayPortalSections(p_portalLayoutType, _
                          p_column_num, _
                          p_orgid, _
                          p_orgRegistration, _
                          p_userid, _
                          p_wrap_td_tags, _
                          p_column_width, _
                          p_showRSS, _
                          p_featurename)

  dim lcl_section_scripts

  lcl_section_scripts = ""

  if p_column_num <> "" then
     lcl_columnNum = cstr(p_column_num)
     lcl_columnNum = dbsafe(lcl_columnNum)
  else
     lcl_columnNum = "1"
  end if

  if p_wrap_td_tags <> "" then
     lcl_wrapTDTags = p_wrap_td_tags
     lcl_wrapTDTags = UCASE(lcl_wrapTDTags)
     lcl_wrapTDTags = dbsafe(lcl_wrapTDTags)
  else
     lcl_wrapTDTags = "Y"
  end if

  if p_column_width <> "" then
     lcl_ColumnWidth = p_column_width
     lcl_ColumnWidth = dbsafe(lcl_ColumnWidth)
  else
     lcl_ColumnWidth = "100%"
  end if

  if p_portalLayoutType <> "" then
     lcl_portalLayoutType = p_portalLayoutType
     lcl_portalLayoutType = UCASE(lcl_portalLayoutType)
     lcl_portalLayoutType = dbsafe(lcl_portalLayoutType)
  else
     lcl_portalLayoutType = "CL"
  end if

  if lcl_portalLayoutType = "SAVVY" then
     lcl_columncheck = "isSavvyOn"
  else
     lcl_columncheck = "isCommunityLinkOn"
  end if

  'if p_showRSS <> "" then
  '   lcl_showRSS = UCASE(p_showRSS)
  'else
  '   lcl_showRSS = "N"
  'end if

  lcl_showRSS = p_showRSS

  if p_featurename <> "" then
     if containsApostrophe(p_featurename) then
        lcl_featurename       = ""
        lcl_query_featurename = ""
     else
        lcl_featureid         = ""
        lcl_featurename       = p_featurename
        lcl_featurename       = ucase(lcl_featurename)
        lcl_query_featurename = lcl_featurename

        lcl_featurename       = dbsafe(lcl_featurename)
        lcl_featurename       = "'" & lcl_featurename & "'"
     end if
  else
     lcl_featurename       = ""
     lcl_query_featurename = ""
  end if

 'Retrieve all of the features for the column specified
  sSQL = " SELECT d.orgid, "
  sSQL = sSQL & " d.featureid, "
  sSQL = sSQL & " d.featurename, "
  sSQL = sSQL & " d.portalcolumn, "
  sSQL = sSQL & " d.displayorder, "
  sSQL = sSQL & " rss_feedid, "
  sSQL = sSQL &   lcl_columncheck & ", "
  sSQL = sSQL & " isnull(d.numListItemsShown_"           & lcl_portalLayoutType & ",1) AS numListItemsShown, "
  sSQL = sSQL & " isnull(d.showsectionborder_"           & lcl_portalLayoutType & ",1) AS showsectionborder, "
  sSQL = sSQL & " isnull(d.sectionbordercolor_"          & lcl_portalLayoutType & ",'000000') AS sectionbordercolor, "
  sSQL = sSQL & " d.sectionbackgroundcolor_"             & lcl_portalLayoutType & " AS sectionbackgroundcolor, "
  sSQL = sSQL & " isnull(d.sectionheader_bgcolor_"       & lcl_portalLayoutType & ",'ffffff') AS sectionheader_bgcolor, "
  sSQL = sSQL & " isnull(d.sectionheader_linecolor_"     & lcl_portalLayoutType & ",'000000') AS sectionheader_linecolor, "
  sSQL = sSQL & " isnull(d.sectionheader_fonttype_"      & lcl_portalLayoutType & ",'" & getCLOptionDefault("SECTIONHEADER_FONTTYPE") & "') AS sectionheader_fonttype, "
  sSQL = sSQL & " isnull(d.sectionheader_fontcolor_"     & lcl_portalLayoutType & ",'000000') AS sectionheader_fontcolor, "
  sSQL = sSQL & " isnull(d.sectionheader_fontsize_"      & lcl_portalLayoutType & ",'11') AS sectionheader_fontsize, "
  sSQL = sSQL & " isnull(d.sectionheader_isbold_"        & lcl_portalLayoutType & ",1) AS sectionheader_isbold, "
  sSQL = sSQL & " isnull(d.sectionheader_isitalic_"      & lcl_portalLayoutType & ",0) AS sectionheader_isitalic, "
  sSQL = sSQL & " isnull(d.sectiontext_bgcolor_"         & lcl_portalLayoutType & ",'ffffff') AS sectiontext_bgcolor, "
  sSQL = sSQL & " isnull(d.sectiontext_bgcolorhover_"    & lcl_portalLayoutType & ",'ffffff') AS sectiontext_bgcolorhover, "
  sSQL = sSQL & " isnull(d.sectiontext_fonttype_"        & lcl_portalLayoutType & ",'" & getCLOptionDefault("SECTIONTEXT_FONTTYPE") & "') AS sectiontext_fonttype, "
  sSQL = sSQL & " isnull(d.sectiontext_fontcolor_"       & lcl_portalLayoutType & ",'000000') AS sectiontext_fontcolor, "
  sSQL = sSQL & " isnull(d.sectiontext_fontcolorhover_"  & lcl_portalLayoutType & ",'000000') AS sectiontext_fontcolorhover, "
  sSQL = sSQL & " isnull(d.sectiontext_fontsize_"        & lcl_portalLayoutType & ",'11') AS sectiontext_fontsize, "
  sSQL = sSQL & " isnull(d.sectionlinks_alignment_"      & lcl_portalLayoutType & ",'" & getCLOptionDefault("SECTIONLINKS_ALIGN")    & "') AS sectionlinks_alignment, "
  sSQL = sSQL & " isnull(d.sectionlinks_fonttype_"       & lcl_portalLayoutType & ",'" & getCLOptionDefault("SECTIONLINKS_FONTTYPE") & "') AS sectionlinks_fonttype, "
  sSQL = sSQL & " isnull(d.sectionlinks_fontcolor_"      & lcl_portalLayoutType & ",'000000') AS sectionlinks_fontcolor, "
  sSQL = sSQL & " isnull(d.sectionlinks_fontcolorhover_" & lcl_portalLayoutType & ",'000000') AS sectionlinks_fontcolorhover, "
  sSQL = sSQL & " isnull(d.viewall_urltype_"             & lcl_portalLayoutType & ",'default') AS viewall_urltype, "
  sSQL = sSQL & " d.viewall_url_"                        & lcl_portalLayoutType & " AS viewall_url, "
  sSQL = sSQL & " isnull(d.viewall_url_wintype_"         & lcl_portalLayoutType & ",'samewindow') AS viewall_url_wintype, "
  sSQL = sSQL & " query_filter "
  sSQL = sSQL & " FROM egov_communitylink_displayorgfeatures d, "
  sSQL = sSQL &      " egov_organizations_to_features FO, "
  sSQL = sSQL &      " egov_organization_features f "
  sSQL = sSQL & " WHERE d.featureid = f.featureid "
  sSQL = sSQL & " AND F.featureid = FO.featureid "
  sSQL = sSQL & " AND FO.orgid = d.orgid "
  sSQL = sSQL & " AND FO.orgid = " & p_orgid
  sSQL = sSQL & " AND d.portalcolumn = " & lcl_ColumnNum
  sSQL = sSQL & " AND " & lcl_columncheck & " = 1 "

  if lcl_featurename <> "" then
     sSQL = sSQL & " AND UPPER(f.feature) = " & lcl_featurename
  end if

  sSQL = sSQL & " ORDER BY isnull(d.displayorder, 1), d.featurename "
'if p_orgid = 153 then
'   dtb_debug p_column_num
'   dtb_debug sSQL
'end if

'response.write sSQL & "<br />"
  set oPortalColumns = Server.CreateObject("ADODB.Recordset")
  oPortalColumns.Open sSQL, Application("DSN"), 3, 1
  if not oPortalColumns.eof then

     if lcl_wrapTDTags = "Y" then
        if lcl_columnNum = 2 then
           lcl_td_bgcolor  = " background-color:#" & oPortalColumns("sectiontext_bgcolor")

           if instr(lcl_ColumnWidth,"%") < 1 then
              lcl_ColumnWidth = lcl_ColumnWidth+5
           end if
        else
           lcl_td_bgcolor = ""
        end if

        response.write "<td height=""100%"" style=""width:" & lcl_ColumnWidth & "px;" & lcl_td_bgcolor & """>" & vbcrlf
     end if

     i = 0
     do while not oPortalColumns.eof
        i = i + 1

       'Determine if the background color for the iFrame BODY tag is overridden for this section.
        lcl_section_backgroundcolor = ""
        lcl_section_scripts         = ""

        if oPortalColumns("sectionbackgroundcolor") <> "" then
           lcl_section_backgroundcolor = "#" & oPortalColumns("sectionbackgroundcolor")
           lcl_section_scripts         = "changeBodyBG('" & lcl_section_backgroundcolor & "');"
        end if

       'Build the Section Header styles
        lcl_sectionheader_styles = ""
        lcl_sectionheader_styles = lcl_sectionheader_styles & "font-family:'" & oPortalColumns("sectionheader_fonttype")  & "';"
        lcl_sectionheader_styles = lcl_sectionheader_styles & "font-size:"    & oPortalColumns("sectionheader_fontsize")  & "px;"
        lcl_sectionheader_styles = lcl_sectionheader_styles & "color:#"       & oPortalColumns("sectionheader_fontcolor") & ";"

        if oPortalColumns("sectionheader_isbold") then
           lcl_sectionheader_styles = lcl_sectionheader_styles & "font-weight:bold;"
        end if

        if oPortalColumns("sectionheader_isitalic") then
           lcl_sectionheader_styles = lcl_sectionheader_styles & "font-style:italic;"
        end if

       'Build the Section Header Layout styles
        lcl_sectionheaderlayout_styles = ""
        lcl_sectionheaderlayout_styles = lcl_sectionheaderlayout_styles & "background-color:#"        & oPortalColumns("sectionheader_bgcolor")   & ";"
        lcl_sectionheaderlayout_styles = lcl_sectionheaderlayout_styles & "border-bottom:1pt solid #" & oPortalColumns("sectionheader_linecolor") & ";"
        lcl_sectionheaderlayout_styles = lcl_sectionheaderlayout_styles & "padding:5px;"

       'Show a top border if this is NOT the 1st section in the column.
        if i > 1 then
           lcl_sectionheaderlayout_styles = lcl_sectionheaderlayout_styles & "border-top:1pt solid #" & oPortalColumns("sectionheader_linecolor") & ";"
        end if

        'if lcl_columnNum = 1 then
           'lcl_sectionheaderlayout_styles = lcl_sectionheaderlayout_styles & "padding-bottom:4px;"
        'end if

       'Build the Section Text styles
        lcl_sectiontext_styles = ""
        lcl_sectiontext_styles = lcl_sectiontext_styles & "font-family:'"      & oPortalColumns("sectiontext_fonttype")  & "';"
        lcl_sectiontext_styles = lcl_sectiontext_styles & "font-size:"         & oPortalColumns("sectiontext_fontsize")  & "px;"
        lcl_sectiontext_styles = lcl_sectiontext_styles & "color:#"            & oPortalColumns("sectiontext_fontcolor") & ";"
        lcl_sectiontext_styles = lcl_sectiontext_styles & "background-color:#" & oPortalColumns("sectiontext_bgcolor")   & ";"
        lcl_sectiontext_styles = lcl_sectiontext_styles & "padding-bottom:2px;"

       'Determine if there is a border around the section.
        lcl_sectionborder_styles = ""

        if oPortalColumns("showsectionborder") then
           if oPortalColumns("sectionbordercolor") <> "" then
              lcl_section_bgcolor = oPortalColumns("sectionbordercolor")
           else
              lcl_section_bgcolor = "000000"
           end if

           lcl_sectionborder_styles = " style=""border:1pt solid #" & lcl_section_bgcolor & ";"""
        end if

        response.write "    <div" & lcl_sectionborder_styles & ">" & vbcrlf
        'response.write "      <div align=""left"" style=""" & lcl_sectionheader_styles & """>" 
        'response.write           oPortalColumns("featurename")
        'response.write "      </div>" & vbcrlf
        response.write "      <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" style=""" & lcl_sectionheaderlayout_styles & """>" & vbcrlf
        response.write "        <tr>" & vbcrlf
        response.write "            <td align=""left"" style=""" & lcl_sectionheader_styles & """>" & oPortalColumns("featurename") & "</td>" & vbcrlf
        response.write "            <td align=""right"">" & vbcrlf

                                     if lcl_showRSS then
                                        'checkForRSSFeed p_orgid, trim(oPortalColumns("featureid")),"", sEgovWebsiteURL
                                        checkForRSSFeed p_orgid, trim(oPortalColumns("rss_feedid")), "", "", sEgovWebsiteURL
                                     end if
        response.write "            </td>" & vbcrlf
        response.write "        </tr>" & vbcrlf
        response.write "      </table>" & vbcrlf
        response.write "      <div align=""left"" style=""" & lcl_sectiontext_styles & """>" & vbcrlf

                               'Check to see if this is a "filtered" feature.  If so, then we want to strip off the "filter"
                               'and get the featureid for the actual feature so we can query the proper results.
                                if instr(lcl_query_featurename,"_filter") > 0 then
                                   lcl_filter_id_start    = instr(lcl_query_featurename,"_filter")
                                   lcl_filter_id          = mid(lcl_query_featurename,lcl_filter_id_start)
                                   lcl_actual_featurename = replace(lcl_query_featurename,lcl_filter_id,"")
                                   lcl_featureid          = getFeatureID(lcl_actual_featurename)
                                else
                                   lcl_featureid = trim(oPortalColumns("featureid"))
                                end if

                               'Get the CL_portaltype
                                lcl_portaltype = getFeaturePortalType(lcl_featureid)

                                getPortalInfo lcl_portalLayoutType, _
                                              p_orgid, _
                                              lcl_featureid, _
                                              p_orgRegistration, _
                                              p_userid, _
                                              oPortalColumns("numListItemsShown"), _
                                              lcl_portaltype, _
                                              oPortalColumns("sectionheader_bgcolor"), _
                                              oPortalColumns("sectiontext_bgcolor"), _
                                              oPortalColumns("sectiontext_bgcolorhover"), _
                                              oPortalColumns("sectiontext_fontcolor"), _
                                              oPortalColumns("sectiontext_fontcolorhover"), _
                                              oPortalColumns("sectiontext_fonttype"), _
                                              oPortalColumns("sectiontext_fontsize"), _
                                              oPortalColumns("sectionlinks_alignment"), _
                                              oPortalColumns("sectionlinks_fonttype"), _
                                              oPortalColumns("sectionlinks_fontcolor"), _
                                              oPortalColumns("sectionlinks_fontcolorhover"), _
                                              oPortalColumns("viewall_urltype"), _
                                              oPortalColumns("viewall_url"), _
                                              oPortalColumns("viewall_url_wintype"), _
                                              oPortalColumns("query_filter")

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

  if lcl_section_scripts <> "" then
     runInlineJavascripts lcl_section_scripts
  end if

end sub

'------------------------------------------------------------------------------
sub getPortalInfo(p_portalLayoutType, _
                  p_orgid, _
                  p_featureid, _
                  p_orgRegistration, _
                  p_userid, _
                  p_numListItemsShown, _
                  p_portaltype, _
                  p_sectionheader_bgcolor, _
                  p_sectiontext_bgcolor, _
                  p_sectiontext_bgcolorhover, _
                  p_sectiontext_fontcolor, _
                  p_sectiontext_fontcolorhover, _
                  p_sectiontext_fonttype, _
                  p_sectiontext_fontsize, _
                  p_sectionlinks_alignment, _
                  p_sectionlinks_fonttype, _
                  p_sectionlinks_fontcolor, _
                  p_sectionlinks_fontcolorhover, _
                  p_viewall_urltype, _
                  p_viewall_url, _
                  p_viewall_url_wintype, _
                  p_query_filter)

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
             displayBlogInfo p_orgid, _
                             p_featureid, _
                             p_orgRegistration, _
                             p_userid, _
                             iNumListItemsShown, _
                             p_sectiontext_fonttype, _
                             p_sectiontext_fontcolor, _
                             p_sectiontext_fontsize, _
                             p_sectionlinks_alignment, _
                             p_sectionlinks_fonttype, _
                             p_sectionlinks_fontcolor, _
                             p_sectionlinks_fontcolorhover, _
                             p_viewall_urltype, _
                             p_viewall_url, _
                             p_viewall_url_wintype, _
                             p_query_filter

          case "COMMUNITY_CALENDAR"
             displayUpcomingEvents p_orgid, _
                                   p_featureid, _
                                   iNumListItemsShown, _
                                   p_sectionheader_bgcolor, _
                                   p_sectiontext_bgcolor, _
                                   p_sectiontext_bgcolorhover, _
                                   p_sectiontext_fonttype, _
                                   p_sectiontext_fontcolor, _
                                   p_sectiontext_fontcolorhover, _
                                   p_sectiontext_fontsize, _
                                   p_sectionlinks_alignment, _
                                   p_sectionlinks_fonttype, _
                                   p_sectionlinks_fontcolor, _
                                   p_sectionlinks_fontcolorhover, _
                                   p_viewall_urltype, _
                                   p_viewall_url, _
                                   p_viewall_url_wintype, _
                                   p_query_filter

          case "FAQ", "RUMORMILL"
             lcl_show_mouseover = "N"

             displayFAQ_RumorMill p_orgid, _
                                  p_featureid, _
                                  p_orgRegistration, _
                                  p_userid, _
                                  iNumListItemsShown, _
                                  p_sectionheader_bgcolor, _
                                  p_sectiontext_bgcolor, _
                                  p_sectiontext_bgcolorhover, _
                                  p_sectiontext_fonttype, _
                                  p_sectiontext_fontcolor, _
                                  p_sectiontext_fontcolorhover, _
                                  p_sectiontext_fontsize, _
                                  p_sectionlinks_alignment, _
                                  p_sectionlinks_fonttype, _
                                  p_sectionlinks_fontcolor, _
                                  p_sectionlinks_fontcolorhover, _
                                  p_viewall_urltype, _
                                  p_viewall_url, _
                                  p_viewall_url_wintype, _
                                  UCASE(iPortalType), _
                                  lcl_show_mouseover, _
                                  p_query_filter

          case "NEWS"
             displayCurrentNews p_orgid, _
                                p_featureid, _
                                iNumListItemsShown, _
                                p_sectionheader_bgcolor, _
                                p_sectiontext_bgcolor, _
                                p_sectiontext_bgcolorhover, _
                                p_sectiontext_fonttype, _
                                p_sectiontext_fontcolor, _
                                p_sectiontext_fontcolorhover, _
                                p_sectiontext_fontsize, _
                                p_sectionlinks_alignment, _
                                p_sectionlinks_fonttype, _
                                p_sectionlinks_fontcolor, _
                                p_sectionlinks_fontcolorhover, _
                                p_viewall_urltype, _
                                p_viewall_url, _
                                p_viewall_url_wintype, _
                                p_query_filter

          case "DOCUMENTS"
             displayNewDocuments p_orgid, _
                                 p_featureid, _
                                 p_userid, _
                                 iNumListItemsShown, _
                                 p_sectionheader_bgcolor, _
                                 p_sectiontext_bgcolor, _
                                 p_sectiontext_bgcolorhover, _
                                 p_sectiontext_fonttype, _
                                 p_sectiontext_fontcolor, _
                                 p_sectiontext_fontcolorhover, _
                                 p_sectiontext_fontsize, _
                                 p_sectionlinks_alignment, _
                                 p_sectionlinks_fonttype, _
                                 p_sectionlinks_fontcolor, _
                                 p_sectionlinks_fontcolorhover, _
                                 p_viewall_urltype, _
                                 p_viewall_url, _
                                 p_viewall_url_wintype, _
                                 p_query_filter
          case else
             response.write "&nbsp;" & vbcrlf
        end select
     else
        response.write "&nbsp;" & vbcrlf
     end if

  end if

end sub

'------------------------------------------------------------------------------
sub displayBlogInfo(p_orgid, _
                    p_featureid, _
                    p_orgRegistration, _
                    p_userid, _
                    p_numListItems, _
                    p_sectiontext_fonttype, _
                    p_sectiontext_fontcolor, _
                    p_sectiontext_fontsize, _
                    p_sectionlinks_alignment, _
                    p_sectionlinks_fonttype, _
                    p_sectionlinks_fontcolor, _
                    p_sectionlinks_fontcolorhover, _
                    p_viewall_urltype, _
                    p_viewall_url, _
                    p_viewall_url_wintype, _
                    p_query_filter)

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

 'Set up the View All url
  lcl_viewAll_url        = sEgovWebsiteURL & "/mayorsblog/mayorsblog.asp"
  lcl_viewAll_openNewWin = "N"

  if p_viewall_urltype <> "" then
     if ucase(p_viewall_urltype) = "CUSTOM" then
        lcl_viewAll_url = p_viewall_url
     end if
  end if

  if p_viewall_url_wintype <> "" then
     if ucase(p_viewall_url_wintype) = "NEWWINDOW" then
        lcl_viewAll_openNewWin = "Y"
     end if
  end if

 'Set up the Post Comments url
  lcl_postcomments_formid = getCommentsFormID(p_orgid, p_featureid, "")
  lcl_postcomments_url    = sEgovWebsiteURL & "/action.asp?actionid=" & lcl_postcomments_formid

  sSQL = "SELECT TOP " & iNumListItems
  sSQL = sSQL & " mb.blogid, "
  sSQL = sSQL & " mb.userid, "
  sSQL = sSQL & " mb.title, "
  sSQL = sSQL & " mb.article, "
  sSQL = sSQL & " mb.createdbyid, "
  sSQL = sSQL & " mb.createdbydate, "
  sSQL = sSQL & " isnull(u.imagefilename,'') AS imagefilename, "
  sSQL = sSQL & " u.firstname + ' ' + u.lastname as createdbyname "
  sSQL = sSQL & " FROM egov_mayorsblog mb, users u "
  sSQL = sSQL & " WHERE mb.userid = u.userid "
  sSQL = sSQL & " AND mb.isInactive = 0 "
  sSQL = sSQL & " AND mb.orgid = " & p_orgid

  if p_query_filter <> "" then
     sSQL = sSQL & p_query_filter
  end if

  sSQL = sSQL & " ORDER BY mb.createdbydate DESC "

  set oBlogInfo = Server.CreateObject("ADODB.Recordset")
  oBlogInfo.Open sSQL, Application("DSN"), 3, 1

  if not oBlogInfo.eof then
     do while not oBlogInfo.eof
        iLineCnt        = iLineCnt + 1
        lcl_display_img = ""
        lcl_img_src     = sEgovWebsiteURL
        lcl_img_border  = "/images/communitylink"
        lcl_img_scripts = ""
        lcl_tooltip     = ""

        if oBlogInfo("imagefilename") <> "" then
           iImgCount         = iImgCount + 1
           lcl_tooltip       = oBlogInfo("createdbyname")
           lcl_imagefilename = oBlogInfo("imagefilename")

           if left(lcl_imagefilename,1) <> "/" then
              lcl_imagefilename = "/" & lcl_imagefilename
           end if

           lcl_blog_imgsrc = ""
           lcl_blog_imgsrc = lcl_blog_imgsrc & Application("CommunityLink_DocUrl")
           lcl_blog_imgsrc = lcl_blog_imgsrc & "/public_documents300/"
           lcl_blog_imgsrc = lcl_blog_imgsrc & sorgVirtualSiteName
           lcl_blog_imgsrc = lcl_blog_imgsrc & "/unpublished_documents"
           lcl_blog_imgsrc = lcl_blog_imgsrc & lcl_imagefilename

           if session("deviceViewMode") <> "M" then
              lcl_display_img = lcl_display_img & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""float:left; margin-right:5px""><br />" & vbcrlf
              lcl_display_img = lcl_display_img & "  <tr><td colspan=""3""><img id=""blogimg_top_" & iImgCount & """ src=""" & lcl_img_src & lcl_img_border & "/blog_img_top.jpg"" height=""16"" alt=""" & lcl_tooltip & """ /></td></tr>" & vbcrlf
              lcl_display_img = lcl_display_img & "  <tr>" & vbcrlf
              lcl_display_img = lcl_display_img & "      <td><img id=""blogimg_left_"  & iImgCount & """ src=""" & lcl_img_src & lcl_img_border & "/blog_img_left.jpg"" width=""11"" alt=""" & lcl_tooltip & """ /></td>" & vbcrlf
              'lcl_display_img = lcl_display_img & "      <td><img id=""blogimg_"       & iImgCount & """ name=""blogimg"" src=""" & lcl_img_src & "/admin/custom/pub/" & sorgVirtualSiteName & "/unpublished_documents" & oBlogInfo("imagefilename") & """ alt=""" & lcl_tooltip & """ /></td>" & vbcrlf
              lcl_display_img = lcl_display_img & "      <td><img id=""blogimg_"       & iImgCount & """ name=""blogimg"" src=""" & lcl_blog_imgsrc & """ alt=""" & lcl_tooltip & """ /></td>" & vbcrlf
              lcl_display_img = lcl_display_img & "      <td><img id=""blogimg_right_" & iImgCount & """ src=""" & lcl_img_src & lcl_img_border & "/blog_img_right.jpg"" width=""15"" alt=""" & lcl_tooltip & """ /></td>" & vbcrlf
              lcl_display_img = lcl_display_img & "  </tr>" & vbcrlf
              lcl_display_img = lcl_display_img & "  <tr><td colspan=""3""><img id=""blogimg_bottom_" & iImgCount & """ src=""" & lcl_img_src & lcl_img_border & "/blog_img_bottom.jpg"" height=""16"" alt=""" & lcl_tooltip & """ /></td></tr>" & vbcrlf
              lcl_display_img = lcl_display_img & "</table>" & vbcrlf
           else
              'lcl_display_img = "<img id=""blogimg_" & iImgCount & """ name=""blogimg"" src=""" & lcl_img_src & "/admin/custom/pub/" & sorgVirtualSiteName & "/unpublished_documents" & oBlogInfo("imagefilename") & """ alt=""" & lcl_tooltip & """ align=""left"" />" & vbcrlf
              lcl_display_img = "<img id=""blogimg_" & iImgCount & """ name=""blogimg"" src=""" & lcl_blog_imgsrc & oBlogInfo("imagefilename") & """ alt=""" & lcl_tooltip & """ align=""left"" />" & vbcrlf
           end if
        else
           iImgCount = iImgCount
        end if

       'Build the View Links URLs
        lcl_viewIndividual_url  = sEgovWebsiteURL & "/mayorsblog/mayorsblog_info.asp?id=" & oBlogInfo("blogid")

       'If there are multiple rows then show the "dotted-line" separator
        if iLineCnt > 1 then
           lcl_section_separator = "border-top:1pt dotted #c0c0c0;"
        else
           lcl_section_separator = ""
        end if

       'Format the Article
        if len(trim(oBlogInfo("article"))) > 500 Then
           lcl_article = StripTags( Trim(oBlogInfo("article")) )
           lcl_article = Left(lcl_article,500) & "..."
           'lcl_article = left(trim(oBlogInfo("article")),500) & "..."
        else
           lcl_article = trim(oBlogInfo("article"))
        end if

        lcl_article = formatArticle(lcl_article)

        response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" style=""" & lcl_section_separator & lcl_section_spacing & """>" & vbcrlf
        response.write "  <tr>" & vbcrlf
        response.write "      <td style=""font-family:" & p_sectiontext_fonttype & "; font-size:" & p_sectiontext_fontsize & "px; color:#" & p_sectiontext_fontcolor & ";"">" & vbcrlf
        response.write            lcl_display_img & vbcrlf
        response.write "          <p>" & vbcrlf
        response.write "             <strong style=""font-size:12px"">" & oBlogInfo("title") & "</strong><br />" & vbcrlf
        response.write "             <i style=""font-size:10px"">by: " & oBlogInfo("createdbyname") & " on " & FormatDateTime(oBlogInfo("createdbydate"),vbshortdate) & "</i>" & vbcrlf
        response.write "          </p>" & vbcrlf
        response.write lcl_article
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "  <tr>" & vbcrlf
        response.write "      <td align=""" & iSectionLinks_Alignment & """ style=""padding:5px;"">" & vbcrlf

       'View Article
        displayViewRowLink "View Article", _
                           "viewIndividual_blog_" & oBlogInfo("blogid"), _
                           lcl_viewIndividual_url, _
                           "changeElementStyles('viewIndividual_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseover_viewlinks & ");", _
                           "changeElementStyles('viewIndividual_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseout_viewlinks  & ");", _
                           "Y"

        response.write "&nbsp;|&nbsp;" & vbcrlf

        lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewIndividual_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

       'View All
        displayViewRowLink "View All", _
                           "viewAll_blog_" & oBlogInfo("blogid"), _
                           lcl_viewAll_url, _
                           "changeElementStyles('viewAll_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseover_viewlinks & ");", _
                           "changeElementStyles('viewAll_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseout_viewlinks  & ");", _
                           lcl_viewAll_openNewWin

        lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_blog_" & oBlogInfo("blogid") & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

       'Post a Comment
       'Only allow the user to enter a comment if a actionline request form has been associated to the blog feature.
        if lcl_orghasfeature_action_line AND lcl_postcomments_formid > 0 then
           response.write "&nbsp;|&nbsp;" & vbcrlf

           displayViewRowLink "Post a Comment", _
                              "postComments_blog_" & oBlogInfo("blogid"), _
                              lcl_postcomments_url, _
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
     displayViewRowLink "View All", _
                        "viewAll_blog_0", _
                        lcl_viewAll_url, _
                        "changeElementStyles('viewAll_blog_0'," & lcl_onmouseover_viewlinks & ");", _
                        "changeElementStyles('viewAll_blog_0'," & lcl_onmouseout_viewlinks  & ");", _
                        "N"

     lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_blog_0'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

     response.write "</div>" & vbcrlf
  end if

  oBlogInfo.close
  set oBlogInfo = nothing

 'Check for any javascripts to run
  runInlineJavascripts lcl_scripts_viewlinks

end sub

'------------------------------------------------------------------------------
sub displayUpcomingEvents(p_orgid, _
                          p_featureid, _
                          p_numListItems, _
                          p_sectionheader_bgcolor, _
                          p_sectiontext_bgcolor, _
                          p_sectiontext_bgcolorhover, _
                          p_sectiontext_fonttype, _
                          p_sectiontext_fontcolor, _
                          p_sectiontext_fontcolorhover, _
                          p_sectiontext_fontsize, _
                          p_sectionlinks_alignment, _
                          p_sectionlinks_fonttype, _
                          p_sectionlinks_fontcolor, _
                          p_sectionlinks_fontcolorhover, _
                          p_viewall_urltype, _
                          p_viewall_url, _
                          p_viewall_url_wintype, _
                          p_query_filter)

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
  lcl_styles_container        = " style=""" & lcl_section_spacing & "cursor:pointer; border-bottom:1pt dotted #c0c0c0; font-family:" & p_sectiontext_fonttype & """"
  'lcl_styles_container_anchor = " style=""text-decoration:none; color:#" & p_sectiontext_bgcolor   & ";"""
  'lcl_styles_eventdate        = " style=""font-size:10px; color:#"       & p_sectiontext_fontcolor & ";"""
  'lcl_styles_subject          = " style=""font-weight:bold; font-size:"  & p_sectiontext_fontsize & "px; color:#" & p_sectiontext_fontcolor & ";"""
  lcl_styles_container_anchor = " style=""text-decoration:none; color:#" & p_sectiontext_fontcolor   & ";"""
  lcl_styles_eventdate        = " style=""font-size:10px;"""
  lcl_styles_subject          = " style=""font-weight:bold; font-size:"  & p_sectiontext_fontsize & "px;"""
  lcl_styles_viewlinks        = " style=""" & lcl_section_spacing & """"

 'Set up the View All url
  lcl_viewAll_url        = sEgovWebsiteURL & "/events/calendar.asp"
  lcl_viewAll_openNewWin = "N"

  if p_viewall_urltype <> "" then
     if ucase(p_viewall_urltype) = "CUSTOM" then
        lcl_viewAll_url = p_viewall_url
     end if
  end if

  if p_viewall_url_wintype <> "" then
     if ucase(p_viewall_url_wintype) = "NEWWINDOW" then
        lcl_viewAll_openNewWin = "Y"
     end if
  end if

  sSQL = "SELECT TOP " & iNumListItems
  sSQL = sSQL & " e.eventid, "
  sSQL = sSQL & " e.eventdate, "
  sSQL = sSQL & " e.subject, "
  sSQL = sSQL & " e.eventduration "
  sSQL = sSQL & " FROM events e "
  sSQL = sSQL & " WHERE e.orgid = " & p_orgid
  sSQL = sSQL & " AND (calendarfeature = '' OR calendarfeature IS NULL) "
  sSQL = sSQL & " AND datediff(dd, '" & Date() & "',e.eventdate) >= 0 "
  sSQL = sSQL & " AND e.isHiddenCL <> 1 "

  if p_query_filter <> "" then
     sSQL = sSQL & p_query_filter
  end if

  sSQL = sSQL & " ORDER BY e.eventdate "

'if p_orgid = 153 then
'dtb_debug sSQL
'end if

'response.write sSQL

  set oUpcomingEvents = Server.CreateObject("ADODB.Recordset")
  oUpcomingEvents.Open sSQL, Application("DSN"), 3, 1

  if not oUpcomingEvents.eof then
     do while not oUpcomingEvents.eof
        iLineCnt = iLineCnt + 1

		'Find the end date
		if oUpcomingEvents("eventduration") > 0 Then
			If CLng(oUpcomingEvents("EventDuration")) = CLng(1440) Then
				dEnd = ""
			Else
				dEnd = dateadd("n",oUpcomingEvents("eventduration"),oUpcomingEvents("eventdate"))

				if datediff("d",dEnd,oUpcomingEvents("eventdate")) = 0 then
					dEnd = FormatDateTime(dEnd,vbLongTime)
				end if

				dEnd = " - " & dEnd
			End If 
		else 
			dEnd = ""
		end if

       'Format the event date/time
        formatEventDateTime oUpcomingEvents("eventdate"),dEnd, sDate1, sDate2

       'Set up the View Links onclick
        lcl_viewIndividual_url  = sEgovWebsiteURL & "/events/calendarevents.asp?date=" & month(oUpcomingEvents("eventdate")) & "-" & day(oUpcomingEvents("eventdate")) & "-" & year(oUpcomingEvents("eventdate"))

        'lcl_onmouseover_event = " onmouseover=""changeElementStyles('events_" & oUpcomingEvents("eventid") & "','','','" & p_sectiontext_fontcolor & "','underline','" & p_sectionheader_bgcolorhover & "');"""
        lcl_onmouseover_event = " onmouseover=""changeElementStyles('events_" & oUpcomingEvents("eventid") & "','" & p_sectiontext_fontsize & "','','" & p_sectiontext_fontcolorhover & "','underline','" & p_sectiontext_bgcolorhover & "');"""
        lcl_onmouseout_event  = " onmouseout=""changeElementStyles('events_"  & oUpcomingEvents("eventid") & "','" & p_sectiontext_fontsize & "','','" & p_sectiontext_fontcolor      & "','none','" & p_sectiontext_bgcolor & "');"""

       'We have to "flip" the <DIV> and <A> tags depending on how the user is accessing the site so that the link works properly.
        if session("deviceViewMode") = "M" then
           response.write "<div id=""events_" & oUpcomingEvents("eventid") & """" & lcl_styles_container & lcl_onmouseover_event & lcl_onmouseout_event & ">" & vbcrlf
           response.write "<a target=""_blank"" href=""" & lcl_viewIndividual_url & """" & lcl_styles_container_anchor & ">" & vbcrlf
        else
           response.write "<a target=""_blank"" href=""" & lcl_viewIndividual_url & """" & lcl_styles_container_anchor & ">" & vbcrlf
           response.write "<div id=""events_" & oUpcomingEvents("eventid") & """" & lcl_styles_container & lcl_onmouseover_event & lcl_onmouseout_event & ">" & vbcrlf
        end if

        'response.write "  <span" & lcl_styles_eventdate & ">" & oUpcomingEvents("eventdate") & "</span><br />" & vbcrlf
        'response.write "  <span" & lcl_styles_eventdate & ">" & MyFormatDateTime(oUpcomingEvents("eventdate"),"&nbsp;") & "</span><br />" & vbcrlf
        response.write "  <span" & lcl_styles_eventdate & ">" & sDate1 & " " & sDate2 & "</span><br />" & vbcrlf
        response.write "  <span" & lcl_styles_subject   & ">" & oUpcomingEvents("subject")                              & "</span>" & vbcrlf

        if session("deviceViewMode") = "M" then
           response.write "</a>" & vbcrlf
           response.write "</div>" & vbcrlf
        else
           response.write "</div>" & vbcrlf
           response.write "</a>" & vbcrlf
        end if

        oUpcomingEvents.movenext
     loop
  end if

  oUpcomingEvents.close
  set oUpcomingEvents = nothing

  response.write "<div align=""" & p_sectionlinks_alignment & """" & lcl_styles_viewlinks & ">" & vbcrlf

 'View All
  displayViewRowLink "View All", _
                     "viewAll_upcomingevents_" & p_featureid, _
                     lcl_viewAll_url, _
                     "changeElementStyles('viewAll_upcomingevents_" & p_featureid & "'," & lcl_onmouseover_viewlinks & ");", _
                     "changeElementStyles('viewAll_upcomingevents_" & p_featureid & "'," & lcl_onmouseout_viewlinks  & ");", _
                     lcl_viewAll_openNewWin

  lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_upcomingevents_" & p_featureid & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

 'Determine if the org has a request form associated with the calendar.
 	sSQL = "SELECT OrgRequestCalOn, "
  sSQL = sSQL & " OrgRequestCalForm "
  sSQL = sSQL & " FROM Organizations "
  sSQL = sSQL & " INNER JOIN TimeZones ON Organizations.OrgTimeZoneID = TimeZones.TimeZoneID "
  sSQL = sSQL & " WHERE orgid = " & p_orgid

 	set oGetOrgInfo = Server.CreateObject("ADODB.Recordset")
	 oGetOrgInfo.Open sSQL, Application("DSN"), 3, 1

  if not oGetOrgInfo.eof then
     lcl_blnCalRequest = oGetOrgInfo("OrgRequestCalOn")
	 	  lcl_iCalForm      = oGetOrgInfo("OrgRequestCalForm")
  else
     lcl_blnCalRequest = False
     lcl_iCalForm      = 0
  end if

  oGetOrgInfo.close
  set oGetOrgInfo = nothing

  if lcl_blnCalRequest then
     response.write "&nbsp;|&nbsp;" & vbcrlf

     lcl_addEventURL = "action.asp?actionid=" & lcl_iCalForm

     displayViewRowLink "Request an Event", _
                        "requestEvent_" & p_featureid, _
                        lcl_addEventURL, _
                        "changeElementStyles('requestEvent_" & p_featureid & "'," & lcl_onmouseover_viewlinks & ");", _
                        "changeElementStyles('requestEvent_" & p_featureid & "'," & lcl_onmouseout_viewlinks  & ");", _
                        "Y"

     lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('requestEvent_" & p_featureid & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf
  end if

  response.write "</div>" & vbcrlf

 'Check for any javascripts to run
  runInlineJavascripts lcl_scripts_viewlinks

end sub

'------------------------------------------------------------------------------
sub displayFAQ_RumorMill(p_orgid, _
                         p_featureid, _
                         p_orgRegistration, _
                         p_userid, _
                         p_numListItems, _
                         p_sectionheader_bgcolor, _
                         p_sectiontext_bgcolor, _
                         p_sectiontext_bgcolorhover, _
                         p_sectiontext_fonttype, _
                         p_sectiontext_fontcolor, _
                         p_sectiontext_fontcolorhover, _
                         p_sectiontext_fontsize, _
                         p_sectionlinks_alignment, _
                         p_sectionlinks_fonttype, _
                         p_sectionlinks_fontcolor, _
                         p_sectionlinks_fontcolorhover, _
                         p_viewall_urltype, _
                         p_viewall_url, _
                         p_viewall_url_wintype, _
                         p_faqtype, _
                         p_showMouseOver, _
                         p_query_filter)

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

'  if p_showMouseOver <> "" then
'     iShowMouseOver = UCASE(p_showMouseOver)
'  else
     iShowMouseOver = "Y"
'  end if

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
  lcl_styles_container        = " style=""" & lcl_section_spacing & "border-bottom:1pt dotted #c0c0c0; font-family:" & p_sectiontext_fonttype & """"
  lcl_styles_container_anchor = " style=""text-decoration:none; color:#" & p_sectiontext_fontcolor   & ";"""
  'lcl_styles_faqRumor         = " style=""font-weight:bold; font-size:" & p_sectiontext_fontsize & "px; color:#" & p_sectiontext_fontcolor & ";"""
  lcl_styles_faqRumor         = " style=""font-weight:bold; font-size:" & p_sectiontext_fontsize & "px;"""
  lcl_styles_viewlinks        = " style=""" & lcl_section_spacing & """"

 'Set up the View All url
  lcl_viewAll_url        = sEgovWebsiteURL & "/faq.asp?faqtype=" & iFAQType
  lcl_viewAll_openNewWin = "N"

  if p_viewall_urltype <> "" then
     if ucase(p_viewall_urltype) = "CUSTOM" then
        lcl_viewAll_url = p_viewall_url
     end if
  end if

  if p_viewall_url_wintype <> "" then
     if ucase(p_viewall_url_wintype) = "NEWWINDOW" then
        lcl_viewAll_openNewWin = "Y"
     end if
  end if

 'Set up the Post Comments urls
  lcl_postcomments_formid = getCommentsFormID(p_orgid, p_featureid, "")
  lcl_postcomments_url    = sEgovWebsiteURL & "/action.asp?actionid=" & lcl_postcomments_formid

  sSQL = "SELECT TOP " & iNumListItems
  sSQL = sSQL & " FAQ.FaqID, "
  sSQL = sSQL & " FAQ.FaqQ, "
  sSQL = sSQL & " FAQ.faqA, "
  sSQL = sSQL & " isnull(faqcategoryname,'') AS faqcategoryname "
  sSQL = sSQL & " FROM FAQ "
 	sSQL = sSQL &      " LEFT OUTER JOIN faq_categories C ON C.faqcategoryid = faq.faqcategoryid "
  sSQL = sSQL &      " AND C.faqtype = faq.faqtype "
 	sSQL = sSQL & " WHERE faq.orgid = " & p_orgid
  sSQL = sSQL & " AND (internalonly = 0 OR internalonly is null) "
  sSQL = sSQL & " AND datediff(dd, isnull(publicationstart,'" & date() & "'), '" & date() & "') >= 0 "
  sSQL = sSQL & " AND datediff(dd, isnull(publicationend,'"   & date() & "'), '" & date() & "') <= 0 "

  if p_query_filter <> "" then
     sSQL = sSQL & p_query_filter
  end if

  sSQL = sSQL & " AND UPPER(faq.faqtype) = '" & iFAQType & "' "

  if iFAQType = "FAQ" then
     sSQL = sSQL & " OR (faq.faqtype = '' OR faq.faqtype IS NULL) "
  end if

 	sSQL = sSQL & " ORDER BY faqid DESC"

  set oFAQRumorInfo = Server.CreateObject("ADODB.Recordset")
  oFAQRumorInfo.Open sSQL, Application("DSN"), 3, 1

  if not oFAQRumorInfo.eof then
     do while not oFAQRumorInfo.eof
        iLineCnt = iLineCnt + 1

       'Set up the View Links onclick
        lcl_viewIndividual_url  = sEgovWebsiteURL & "/faq_info.asp?faqtype=" & iFAQType & "&id=" & oFAQRumorInfo("faqid")

        'if iShowMouseOver = "Y" then
           'lcl_onmouseover_faqRumor = " onmouseover=""changeElementStyles('" & iFAQType & "_" & oFAQRumorInfo("faqid") & "','" & p_sectiontext_fontsize & "','','" & p_sectiontext_fontcolor & "','underline','" & p_sectiontext_bgcolorhover & "');"""
           'lcl_onmouseout_faqRumor  = " onmouseout=""changeElementStyles('"  & iFAQType & "_" & oFAQRumorInfo("faqid") & "','" & p_sectiontext_fontsize & "','','','none','" & p_sectiontext_bgcolor & "');"""
           lcl_onmouseover_faqRumor = " onmouseover=""changeElementStyles('" & iFAQType & "_" & oFAQRumorInfo("faqid") & "','" & p_sectiontext_fontsize & "','','" & p_sectiontext_fontcolorhover & "','underline','" & p_sectiontext_bgcolorhover & "');"""
           lcl_onmouseout_faqRumor  = " onmouseout=""changeElementStyles('"  & iFAQType & "_" & oFAQRumorInfo("faqid") & "','" & p_sectiontext_fontsize & "','','" & p_sectiontext_fontcolor      & "','none','" & p_sectiontext_bgcolor & "');"""
        'end if

       'We have to "flip" the <DIV> and <A> tags depending on how the user is accessing the site so that the link works properly.
        if session("deviceViewMode") = "M" then
           response.write "<div id=""" & iFAQType & "_" & oFAQRumorInfo("faqid") & """" & lcl_styles_container & lcl_onmouseover_faqRumor & lcl_onmouseout_faqRumor & ">" & vbcrlf
           response.write "<a target=""_blank"" href=""" & lcl_viewIndividual_url & """" & lcl_styles_container_anchor & ">" & vbcrlf
        else
           response.write "<a target=""_blank"" href=""" & lcl_viewIndividual_url & """" & lcl_styles_container_anchor & ">" & vbcrlf
           response.write "<div id=""" & iFAQType & "_" & oFAQRumorInfo("faqid") & """" & lcl_styles_container & lcl_onmouseover_faqRumor & lcl_onmouseout_faqRumor & ">" & vbcrlf
        end if

        'response.write "  <li><span" & lcl_styles_faqRumor & ">" & oFAQRumorInfo("FaqQ") & "</span>" & vbcrlf
        response.write "  <span" & lcl_styles_faqRumor & ">&bull;&nbsp;" & oFAQRumorInfo("FaqQ") & "</span>" & vbcrlf

        if session("deviceViewMode") = "M" then
           response.write "</a>" & vbcrlf
           response.write "</div>" & vbcrlf
        else
           response.write "</div>" & vbcrlf
           response.write "</a>" & vbcrlf
        end if

        oFAQRumorInfo.movenext
     loop
  end if

  oFAQRumorInfo.close
  set oFAQRumorInfo = nothing

  response.write "<div align=""" & p_sectionlinks_alignment & """" & lcl_styles_viewlinks & ">" & vbcrlf

 'View All
  displayViewRowLink "View All", _
                     "viewAll_" & iFAQType & "_" & p_featureid, _
                     lcl_viewAll_url, _
                     "changeElementStyles('viewAll_" & iFAQType & "_" & p_featureid & "'," & lcl_onmouseover_viewlinks & ");", _
                     "changeElementStyles('viewAll_" & iFAQType & "_" & p_featureid & "'," & lcl_onmouseout_viewlinks  & ");", _
                     lcl_viewAll_openNewWin

  lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_" & iFAQType & "_" & p_featureid & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

 'Submit a Rumor/Ask a Question (Post a Comment)
 'Only allow the user to enter a comment if the feature has a request associated to it..
 ' - p_orgRegistration = sOrgRegistration (found in classOrganization)
  if lcl_orghasfeature_action_line AND lcl_postcomments_formid > 0 then
     'if iFAQType = "FAQ" then
     '   lcl_comments_label = "Ask a Question"
     'else
     '   lcl_comments_label = "Submit a Rumor"
     'end if

     lcl_comments_label = ""
     lcl_comments_label = getCommentsLabel(p_orgid, p_featureid, "")

     if lcl_comments_label = "" OR isnull(lcl_comments_label) then
        lcl_comments_label = "Ask a Question"
     end if

     response.write "&nbsp;|&nbsp;" & vbcrlf

     displayViewRowLink lcl_comments_label, _
                        "postComments_" & iFAQType & "_" & p_featureid, _
                        lcl_postcomments_url, _
                        "changeElementStyles('postComments_" & iFAQType & "_" & p_featureid & "'," & lcl_onmouseover_viewlinks & ");", _
                        "changeElementStyles('postComments_" & iFAQType & "_" & p_featureid & "'," & lcl_onmouseout_viewlinks  & ");", _
                        "N"

     lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('postComments_" & iFAQType & "_" & p_featureid & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf
  end if

  response.write "</div>" & vbcrlf

 'Check for any javascripts to run
  runInlineJavascripts lcl_scripts_viewlinks

end sub

'------------------------------------------------------------------------------
sub displayCurrentNews(p_orgid, _
                       p_featureid, _
                       p_numListItems, _
                       p_sectionheader_bgcolor, _
                       p_sectiontext_bgcolor, _
                       p_sectiontext_bgcolorhover, _
                       p_sectiontext_fonttype, _
                       p_sectiontext_fontcolor, _
                       p_sectiontext_fontcolorhover, _
                       p_sectiontext_fontsize, _
                       p_sectionlinks_alignment, _
                       p_sectionlinks_fonttype, _
                       p_sectionlinks_fontcolor, _
                       p_sectionlinks_fontcolorhover, _
                       p_viewall_urltype, _
                       p_viewall_url, _
                       p_viewall_url_wintype, _
                       p_query_filter)

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
  lcl_section_spacing         = "padding:5px 10px 5px 10px;"
  lcl_styles_container        = " style=""" & lcl_section_spacing & "cursor:pointer; border-bottom:1pt dotted #c0c0c0; font-family:" & p_sectiontext_fonttype & """"
  'lcl_styles_container_anchor = " style=""text-decoration:none; color:#" & p_sectiontext_bgcolor   & ";"""
  'lcl_styles_articledate      = " style=""font-size:10px; color:#"       & p_sectiontext_fontcolor & ";"""
  'lcl_styles_itemtitle        = " style=""font-weight:bold; font-size:"  & p_sectiontext_fontsize  & "px; color:#" & p_sectiontext_fontcolor & ";"""
  lcl_styles_container_anchor = " style=""text-decoration:none; color:#" & p_sectiontext_fontcolor   & ";"""
  lcl_styles_articledate      = " style=""font-size:10px;"""
  lcl_styles_itemtitle        = " style=""font-weight:bold; font-size:"  & p_sectiontext_fontsize  & "px;"""
  lcl_styles_viewlinks        = " style=""" & lcl_section_spacing & """"

 'Set up the View All url
  lcl_viewAll_url        = sEgovWebsiteURL & "/news/news.asp"
  lcl_viewAll_openNewWin = "N"

  if p_viewall_urltype <> "" then
     if ucase(p_viewall_urltype) = "CUSTOM" then
        lcl_viewAll_url = p_viewall_url
     end if
  end if

  if p_viewall_url_wintype <> "" then
     if ucase(p_viewall_url_wintype) = "NEWWINDOW" then
        lcl_viewAll_openNewWin = "Y"
     end if
  end if

 'Set up the Suggest a News Item urls
  lcl_postcomments_formid = getCommentsFormID(p_orgid, p_featureid, "")
  lcl_postcomments_url    = sEgovWebsiteURL & "/action.asp?actionid=" & lcl_postcomments_formid

  sSQL = "SELECT TOP " & iNumListItems
  sSQL = sSQL & " newsitemid, "
  sSQL = sSQL & " itemtitle, "
  sSQL = sSQL & " isnull(publicationstart,itemdate) AS articledate "
  sSQL = sSQL & " FROM egov_news_items "
  sSQL = sSQL & " WHERE itemdisplay = 1 "
  sSQL = sSQL & " AND orgid = "  & p_orgid
  sSQL = sSQL & " AND datediff(dd, '" & date() & "',isnull(publicationstart,itemdate)) <= 0 "
  sSQL = sSQL & " AND (publicationend IS NULL OR datediff(dd, '" & date() & "',isnull(publicationend,'" & date() & "')) >= 0) "
  sSQL = sSQL & " AND UPPER(newstype) = 'NEWS' "

  if p_query_filter <> "" then
     sSQL = sSQL & p_query_filter
  end if

  sSQL = sSQL & " ORDER BY isnull(publicationstart,itemdate) DESC "

  set oCurrentNews = Server.CreateObject("ADODB.Recordset")
  oCurrentNews.Open sSQL, Application("DSN"), 3, 1

  if not oCurrentNews.eof then
     do while not oCurrentNews.eof
        iLineCnt = iLineCnt + 1

       'Set up the View Links onclick
        lcl_viewIndividual_url  = sEgovWebsiteURL & "/news/news_info.asp?id=" & oCurrentNews("newsitemid")
        'lcl_onmouseover_event = " onmouseover=""changeElementStyles('news_" & oCurrentNews("newsitemid") & "','" & p_sectiontext_fontsize & "','','" & p_sectiontext_fontcolor & "','underline','" & p_sectiontext_bgcolorhover & "');"""
        'lcl_onmouseout_event  = " onmouseout=""changeElementStyles('news_"  & oCurrentNews("newsitemid") & "','" & p_sectiontext_fontsize & "','','','none','" & p_sectiontext_bgcolor & "');"""
        lcl_onmouseover_event = " onmouseover=""changeElementStyles('news_" & oCurrentNews("newsitemid") & "','" & p_sectiontext_fontsize & "','','" & p_sectiontext_fontcolorhover & "','underline','" & p_sectiontext_bgcolorhover & "');"""
        lcl_onmouseout_event  = " onmouseout=""changeElementStyles('news_"  & oCurrentNews("newsitemid") & "','" & p_sectiontext_fontsize & "','','" & p_sectiontext_fontcolor      & "','none','" & p_sectiontext_bgcolor & "');"""

       'We have to "flip" the <DIV> and <A> tags depending on how the user is accessing the site so that the link works properly.
        if session("deviceViewMode") = "M" then
           response.write "<div id=""news_" & oCurrentNews("newsitemid") & """" & lcl_styles_container & lcl_onmouseover_event & lcl_onmouseout_event & ">" & vbcrlf
           response.write "<a href=""javascript:openWin('" & lcl_viewIndividual_url & "','','news',400,250);""" & lcl_styles_container_anchor & ">" & vbcrlf
        else
           response.write "<a href=""javascript:openWin('" & lcl_viewIndividual_url & "','','news',400,250);""" & lcl_styles_container_anchor & ">" & vbcrlf
           response.write "<div id=""news_" & oCurrentNews("newsitemid") & """" & lcl_styles_container & lcl_onmouseover_event & lcl_onmouseout_event & ">" & vbcrlf
        end if

        response.write "  <span" & lcl_styles_articledate & ">" & oCurrentNews("articledate") & "</span><br />" & vbcrlf
        response.write "  <span" & lcl_styles_itemtitle   & ">" & oCurrentNews("itemtitle")   & "</span>" & vbcrlf

        if session("deviceViewMode") = "M" then
           response.write "</a>" & vbcrlf
           response.write "</div>" & vbcrlf
        else
           response.write "</div>" & vbcrlf
           response.write "</a>" & vbcrlf
        end if

        oCurrentNews.movenext
     loop

  end if

  oCurrentNews.close
  set oCurrentNews = nothing

  response.write "<div align=""" & p_sectionlinks_alignment & """" & lcl_styles_viewlinks & ">" & vbcrlf

 'View All
  displayViewRowLink "View All", _
                     "viewAll_currentNews_" & p_featureid, _
                     lcl_viewAll_url, _
                     "changeElementStyles('viewAll_currentNews_" & p_featureid & "'," & lcl_onmouseover_viewlinks & ");", _
                     "changeElementStyles('viewAll_currentNews_" & p_featureid & "'," & lcl_onmouseout_viewlinks  & ");", _
                     lcl_viewAll_openNewWin

  lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_currentNews_" & p_featureid & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

 'Post a Comment
 'Only allow the user to enter a comment if a actionline request form has been associated to the blog feature.
  if lcl_orghasfeature_action_line AND lcl_postcomments_formid > 0 then
     response.write "&nbsp;|&nbsp;" & vbcrlf

     displayViewRowLink "Suggest a News Item", _
                        "postComments_news_" & p_featureid, _
                        lcl_postcomments_url, _
                        "changeElementStyles('postComments_news_" & p_featureid & "'," & lcl_onmouseover_viewlinks & ");", _
                        "changeElementStyles('postComments_news_" & p_featureid & "'," & lcl_onmouseout_viewlinks  & ");", _
                        "N"

     lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('postComments_news_" & p_featureid & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf
  end if

  response.write "</div>" & vbcrlf

 'Check for any javascripts to run
  runInlineJavascripts lcl_scripts_viewlinks

end sub

'------------------------------------------------------------------------------
sub displayNewDocuments(p_orgid, _
                        p_featureid, _
                        p_userid, _
                        p_numListItems, _
                        p_sectionheader_bgcolor, _
                        p_sectiontext_bgcolor, _
                        p_sectiontext_bgcolorhover, _
                        p_sectiontext_fonttype, _
                        p_sectiontext_fontcolor, _
                        p_sectiontext_fontcolorhover, _
                        p_sectiontext_fontsize, _
                        p_sectionlinks_alignment, _
                        p_sectionlinks_fonttype, _
                        p_sectionlinks_fontcolor, _
                        p_sectionlinks_fontcolorhover, _
                        p_viewall_urltype, _
                        p_viewall_url, _
                        p_viewall_url_wintype, _
                        p_query_filter)

  iLineCnt              = 1
  iCountRestricted      = 0
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
  lcl_styles_container        = " style=""" & lcl_section_spacing & "border-bottom:1pt dotted #c0c0c0; font-family:" & p_sectiontext_fonttype & """"
  'lcl_styles_container_anchor = " style=""text-decoration:none; color:#" & p_sectiontext_fontcolor & ";"""
  'lcl_styles_newDoc           = " style=""font-weight:bold; font-size:" & p_sectiontext_fontsize & "px; color:#" & p_sectiontext_fontcolor & ";"""
  lcl_styles_container_anchor = " style=""color:#" & p_sectiontext_fontcolor & ";"""
  lcl_styles_newDoc           = " style=""font-weight:bold; font-size:" & p_sectiontext_fontsize & "px;"""
  lcl_styles_viewlinks        = " style=""" & lcl_section_spacing & """"

 'Set up the View All url
  lcl_viewAll_url        = sEgovWebsiteURL & "/docs/menu/home.asp"
  lcl_viewAll_openNewWin = "N"

  if p_viewall_urltype <> "" then
     if ucase(p_viewall_urltype) = "CUSTOM" then
        lcl_viewAll_url = p_viewall_url
     end if
  end if

  if p_viewall_url_wintype <> "" then
     if ucase(p_viewall_url_wintype) = "NEWWINDOW" then
        lcl_viewAll_openNewWin = "Y"
     end if
  end if

 'Get the total document count for the org
  lcl_totaldocs = 0

  sSQL = "SELECT count(d.documentid) as total_docs "
  sSQL = sSQL & " FROM documents d "
  sSQL = sSQL & " WHERE d.orgid = " & p_orgid
  sSQL = sSQL & " AND UPPER(d.documenturl) LIKE ('%/PUBLISHED_DOCUMENTS%') "

  set oDocCount = Server.CreateObject("ADODB.Recordset")
  oDocCount.Open sSQL, Application("DSN"), 3, 1

  if not oDocCount.eof then
     lcl_totaldocs = oDocCount("total_docs")
  end if

  oDocCount.close
  set oDocCount = nothing

 'Get the documents
		'sSQL = "SELECT TOP " & iNumListItems & " d.documentid, d.documenturl, d.documenttitle, d.dateadded, d.parentfolderid, "
  sSQL = "SELECT d.documentid, "
  sSQL = sSQL & " d.documenturl, "
  sSQL = sSQL & " d.documenttitle, "
  sSQL = sSQL & " d.dateadded, "
  sSQL = sSQL & " d.parentfolderid, "
  sSQL = sSQL & " (select df.folderpath "
  sSQL = sSQL &  " from documentfolders df "
  sSQL = sSQL &  " where df.folderid = d.parentfolderid) AS parentfolderurl "
  sSQL = sSQL & " FROM documents d "
  sSQL = sSQL & " WHERE d.orgid = " & p_orgid
  sSQL = sSQL & " AND UPPER(d.documenturl) LIKE ('%/PUBLISHED_DOCUMENTS%') "

  if p_query_filter <> "" then
     sSQL = sSQL & p_query_filter
  end if

  sSQL = sSQL & " ORDER BY d.dateadded desc "

  set oNewDocs = Server.CreateObject("ADODB.Recordset")
  oNewDocs.Open sSQL, Application("DSN"), 3, 1

  if not oNewDocs.eof then

     iTotalCnt = 0

     'do while not oNewDocs.eof
     do until iLineCnt > iNumListItems
        iTotalCnt = iTotalCnt + 1
        iShowDoc  = 0

       'This check will verify that the document exists in the location within the filesystem.
        lcl_valid_doc = checkDocExists(oNewDocs("documenturl"))

        if lcl_valid_doc then
          'Determine if the any of the documents are restricted.
           lcl_restrictdocs_exist = checkDocRestrictionExists(p_orgid, oNewDocs("parentfolderurl"))

          'If restricted documents exists now verify that the the user has access to see them.
          'If no restricted documents exist then simply show the documents
           if lcl_restrictdocs_exist then
              lcl_hasaccess = checkDocumentAccess(p_orgid,p_userid,oNewDocs("parentfolderurl"))

            'Track the number of "restricted" documents were not shown.
              if lcl_hasaccess then
                 iShowDoc         = 1
                 iCountRestricted = iCountRestricted
              else
                 iShowDoc         = 0
                 iCountRestricted = iCountRestricted + 1
              end if
           else
              iShowDoc = 1
           end if

          'Display the document if there are no restrictions
           if iShowDoc = 1 then
              iLineCnt = iLineCnt + 1
           
             'Set up the View Links onclick
              'lcl_viewIndividual_url  = sEgovWebsiteURL & "/admin"
              'lcl_viewIndividual_url = lcl_viewIndividual_url & replace(oNewDocs("documenturl"),"/public_documents300","")

              lcl_viewIndividual_url = ""
              lcl_viewIndividual_url = lcl_viewIndividual_url & Application("CommunityLink_DocUrl")
              'lcl_viewIndividual_url = lcl_viewIndividual_url & "public_documents300/"
              'lcl_viewIndividual_url = lcl_viewIndividual_url & sorgVirtualSiteName
              'lcl_viewIndividual_url = lcl_viewIndividual_url & replace(replace(oNewDocs("documenturl"),"/public_documents300",""),"/custom/pub/","")
              lcl_viewIndividual_url = lcl_viewIndividual_url & oNewDocs("documenturl")
              lcl_viewIndividual_url = replace(lcl_viewIndividual_url,"/custom/pub/","")
              lcl_viewIndividual_url = replace(lcl_viewIndividual_url,"/public_documents300" & sorgVirtualSiteName & "/","public_documents300/" & sorgVirtualSiteName & "/")


              'lcl_onmouseover_newDoc = " onmouseover=""changeElementStyles('newdoc_" & oNewDocs("documentid") & "','" & p_sectiontext_fontsize & "','','" & p_sectiontext_fontcolor & "','underline','" & p_sectiontext_bgcolorhover & "');"""
              'lcl_onmouseout_newDoc  = " onmouseout=""changeElementStyles('newdoc_"  & oNewDocs("documentid") & "','" & p_sectiontext_fontsize & "','','','none','" & p_sectiontext_bgcolor & "');"""
              lcl_onmouseover_newDoc = " onmouseover=""changeElementStyles('newdoc_" & oNewDocs("documentid") & "','" & p_sectiontext_fontsize & "','','" & p_sectiontext_fontcolorhover & "','underline','" & p_sectiontext_bgcolorhover & "');"""
              lcl_onmouseout_newDoc  = " onmouseout=""changeElementStyles('newdoc_"  & oNewDocs("documentid") & "','" & p_sectiontext_fontsize & "','','" & p_sectiontext_fontcolor      & "','none','" & p_sectiontext_bgcolor & "');"""

             'We have to "flip" the <DIV> and <A> tags depending on how the user is accessing the site so that the link works properly.
              if session("deviceViewMode") = "M" then
                 response.write "<div id=""newdoc_" & oNewDocs("documentid") & """" & lcl_styles_container & lcl_onmouseover_newDoc & lcl_onmouseout_newDoc & ">" & vbcrlf
                 response.write "<a target=""_blank"" href=""" & lcl_viewIndividual_url & """" & lcl_styles_container_anchor & ">" & vbcrlf
              else
                 response.write "<a target=""_blank"" href=""" & lcl_viewIndividual_url & """" & lcl_styles_container_anchor & ">" & vbcrlf
                 response.write "<div id=""newdoc_" & oNewDocs("documentid") & """" & lcl_styles_container & lcl_onmouseover_newDoc & lcl_onmouseout_newDoc & ">" & vbcrlf
              end if

              'response.write "  <li style=""padding-bottom:5px"">" & vbcrlf
              response.write "    <span" & lcl_styles_newDoc & ">&bull;&nbsp;" & oNewDocs("documenttitle") & "&nbsp;</span>" & vbcrlf
              'response.write "    <a href=""" & sEgovWebsiteURL & "/docs/menu/home.asp?path=" & oNewDocs("documenturl") & """ style=""position:absolute; z-index:1;"">" & vbcrlf
              'response.write "    <img src=""" & sEgovWebsiteURL & "/images/communitylink/docfolder.png"" width=""16"" height=""13"" border=""0"" align=""top"" alt=""Click to open folder"" style=""position:absolute; float:right;"" />" & vbcrlf
              'response.write "    </a>" & vbcrlf
              response.write "    <input type=""image"" src=""" & sEgovWebsiteURL & "/images/communitylink/docfolder.png"" value=""test"" alt=""Open Folder"" onclick=""window.open('" & sEgovWebsiteURL & "/docs/menu/home.asp?path=" & oNewDocs("documenturl") & "','_top');"" style=""position:absolute; float:right;"" />" & vbcrlf
              'response.write "    <input type=""image"" src=""" & sEgovWebsiteURL & "/images/communitylink/docfolder.png"" value=""test"" alt=""Open Folder"" onclick=""location.href='" & sEgovWebsiteURL & "/docs/menu/home.asp?path=" & oNewDocs("documenturl") & "';"" style=""position:absolute; float:right;"" />" & vbcrlf
              'response.write "  </li>" & vbcrlf

              if session("deviceViewMode") = "M" then
                 response.write "</a>" & vbcrlf
                 response.write "</div>" & vbcrlf
              else
                 response.write "</div>" & vbcrlf
                 response.write "</a>" & vbcrlf
              end if
           else
              iLineCnt = iLineCnt
           end if
        end if

        if iTotalCnt = lcl_totaldocs then
           exit do
        else
           oNewDocs.movenext
        end if
     loop
  end if

  oNewDocs.close
  set oNewDocs = nothing

 'Display "no documents found" message
  if iLineCnt = 0 then
     if iCountRestricted > 0 then
        lcl_nodisplay_message = "All recently added documents have been restricted.  You must be granted proper access to view them."
     else
        lcl_nodisplay_message = "No documents have been recently added."
     end if

     response.write "<div" & lcl_styles_container & "><span style=""color:#800000"">" & lcl_nodisplay_message & "</span><br /><br /></div>" & vbcrlf

  end if
  
  response.write "<div align=""" & p_sectionlinks_alignment & """" & lcl_styles_viewlinks & ">" & vbcrlf

 'View All
  displayViewRowLink "View All", _
                     "viewAll_newDoc_" & p_featureid, _
                     lcl_viewAll_url, _
                     "changeElementStyles('viewAll_newDoc_" & p_featureid & "'," & lcl_onmouseover_viewlinks & ");", _
                     "changeElementStyles('viewAll_newDoc_" & p_featureid & "'," & lcl_onmouseout_viewlinks  & ");", _
                     lcl_viewAll_openNewWin

  lcl_scripts_viewlinks = lcl_scripts_viewlinks & "changeElementStyles('viewAll_newDoc_" & p_featureid & "'," & lcl_onmouseout_viewlinks & ");" & vbcrlf

  response.write "</div>" & vbcrlf

 'Check for any javascripts to run
  runInlineJavascripts lcl_scripts_viewlinks

end sub

'-- Copied from "public-site"/docs/menu/home.asp ------------------------------
function checkDocumentAccess(p_orgid, _
                             p_userid, _
                             p_filepath)

  iReturnValue = False

 'Determine if the user is logged in or not
  if p_userid <> "" AND p_userid <> "-1" then
     lcl_userid = p_userid
  else
     lcl_userid = 0
  end if

  On Error Resume Next

  sSQL = "EXEC CHECKFOLDERACCESS '" & p_orgid & "','" & lcl_userid & "','" & p_filepath & "'"

  set oHasAccess = Server.CreateObject("ADODB.Recordset")
  oHasAccess.Open sSQL, Application("DSN"), 3, 1

  if not oHasAccess.eof then
   		'if oHasAccess("folderid") >= 0 then
   		if oHasAccess("folderid") > 0 then
      		iReturnValue = True
     end if
  end if

  oHasAccess.close
  set oHasAccess = nothing

  checkDocumentAccess = iReturnValue

end function

'------------------------------------------------------------------------------
function checkDocRestrictionExists(p_orgid, _
                                   p_filepath)
  lcl_return   = False
  lcl_folderid = 0

  if p_filepath <> "" then
    'Get the folder id for the file path
     sSQL = "SELECT folderid "
     sSQL = sSQL & " FROM DocumentFolders "
     sSQL = sSQL & " WHERE UPPER(folderpath) = '" & UCASE(p_filepath) & "'"

     set oGetFolderID = Server.CreateObject("ADODB.Recordset")
     oGetFolderID.Open sSQL, Application("DSN"), 3, 1

     if not oGetFolderID.eof then
        lcl_folderid = oGetFolderID("folderid")
     end if

     oGetFolderID.close
     set oGetFolderID = nothing

    'If a folder id exists then check to see if the org has restricted the folder.
     if lcl_folderid > 0 then
        sSQL = "SELECT distinct df.folderid, df.foldername, df.folderpath, isnull(df.CitizenAccessID,0) [Secure] "
        sSQL = sSQL & " FROM documentfolders df "
        sSQL = sSQL & " LEFT JOIN CitizenFeatureAccess fa ON fa.accessid = df.citizenaccessid "
        sSQL = sSQL & " WHERE df.folderid = " & lcl_folderid
        sSQL = sSQL & " AND df.orgid = " & p_orgid

        set oDocsRestrictExists = Server.CreateObject("ADODB.Recordset")
        oDocsRestrictExists.Open sSQL, Application("DSN"), 3, 1

        if not oDocsRestrictExists.eof then
           'lcl_return = oDocsRestrictExists("secure")
           if oDocsRestrictExists("secure") > 0 then
              lcl_return = true
           end if
        end if

        oDocsRestrictExists.close
        set oDocsRestrictExists = nothing
     end if
  end if

  checkDocRestrictionExists = lcl_return

end function

'------------------------------------------------------------------------------
'Function GetVirtualDirectyName()  'Copied from include_top_functions.asp
'	sReturnValue = ""
	
'	strURL = Request.ServerVariables("SCRIPT_NAME")
'	strURL = Split(strURL, "/", -1, 1) 
'	sReturnValue = "/" & strURL(1) 

'	GetVirtualDirectyName = replace(sReturnValue,"/","")

'End Function

'------------------------------------------------------------------------------
function checkDocExists(sPath)
  lcl_return = False

  if sPath <> "" then

     set fs = Server.CreateObject("Scripting.FileSystemObject")

	on error resume next
	mappedpath = server.mappath(replace(sPath,"/custom/pub",""))
	errnum = err.number
	on error goto 0

     if fs.FileExists(mappedpath) = true then
        lcl_return = True
     end if

     set fs = nothing

	 if errnum <> 0 then lcl_return = false
  end if

  checkDocExists = lcl_return

end function

'------------------------------------------------------------------------------
function getTotalColumns(p_orgid)

  lcl_return = 1

 'Find a distinct count of columns to display.
 'NOTE: there MUST always be a return of 1 to build the page properly. (0 = 1)
  sSQL = "SELECT count(distinct portalcolumn) as total_columns "
  sSQL = sSQL & " FROM egov_communitylink_displayorgfeatures "
  sSQL = sSQL & " WHERE orgid = " & p_orgid

  set oGetTotalColumns = Server.CreateObject("ADODB.Recordset")
  oGetTotalColumns.Open sSQL, Application("DSN"), 3, 1

  if not oGetTotalColumns.eof then
     if oGetTotalColumns("total_columns") = 0 then
        lcl_return = 1
     else
        lcl_return = oGetTotalColumns("total_columns")
     end if
  end if

  oGetTotalColumns.close
  set oGetTotalColumns = nothing

  getTotalColumns = lcl_return

end function

'------------------------------------------------------------------------------
function dbsafe(iValue)
  dim lcl_return

  lcl_return = ""

  if iValue <> "" then
     lcl_return = iValue
     lcl_return = replace(lcl_return,"'","''")
  end if

  dbsafe = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  dim sSQL, oDTB

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES('" & replace(p_value,"'","''") & "')"
  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1

  set oDTB = nothing
end sub
%>
