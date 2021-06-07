<!-- #include file="../includes/common.asp" //-->
<!-- #include file="communitylink_global_functions.asp" //-->
<%
Call updateCommunityLink(request("communitylinkid"), _
                         request("isegovhomepage"), _
                         request("website_size"), _
                         request("website_size_customsize"), _
                         request("website_alignment"), _
                         request("website_bgcolor"), _
                         request("showlogo"), _
                         request("logo_filename"), _
                         request("logo_filenamebg"), _
                         request("logo_alignment"), _
                         request("showtopbar"), _
                         request("topbar_bgcolor"), _
                         request("topbar_fonttype"), _
                         request("topbar_fontcolor"), _
                         request("topbar_fontcolorhover"), _
                         request("showsidemenubar"), _
                         request("sidemenubar_alignment"), _
                         request("sidemenuoption_bgcolor"), _
                         request("sidemenuoption_bgcolorhover"), _
                         request("sidemenuoption_alignment"), _
                         request("sidemenuoption_fonttype"), _
                         request("sidemenuoption_fontcolor"), _
                         request("sidemenuoption_fontcolorhover"), _
                         request("showpageheader"), _
                         request("pageheader_alignment"), _
                         request("pageheader_fontsize"), _
                         request("pageheader_fontcolor"), _
                         request("pageheader_fonttype"), _
                         request("pageheader_bgcolor"), _
                         request("showfooter"), _
                         request("footer_bgcolor"), _
                         request("footer_fonttype"), _
                         request("footer_fontcolor"), _
                         request("footer_fontcolorhover"), _
                         request("totalCLRows"), _
                         request("showRSS"), _
                         request("url_twitter"), _
                         request("url_facebook"), _
                         request("url_myspace"), _
                         request("url_blogger"))

'------------------------------------------------------------------------------
sub updateCommunityLink(iCommunityLinkID, _
                        iIsEgovHomePage, _
                        iWebsite_Size, _
                        iWebsite_Size_CustomSize, _
                        iWebsite_Alignment, _
                        iWebsite_BGColor, _
                        iShowLogo, _
                        iLogo_Filename, _
                        iLogo_FilenameBG, _
                        iLogo_Alignment, _
                        iShowTopBar, _
                        iTopBar_BGColor, _
                        iTopBar_FontType, _
                        iTopBar_FontColor, _
                        iTopBar_FontColorHover, _
                        iShowSideMenuBar, _
                        iSideMenuBar_Alignment, _
                        iSideMenuOption_BGColor, _
                        iSideMenuOption_BGColorHover, _
                        iSideMenuOption_Alignment, _
                        iSideMenuOption_FontType, _
                        iSideMenuOption_FontColor, _
                        iSideMenuOption_FontColorHover, _
                        iShowPageHeader, _
                        iPageHeader_Alignment, _
                        iPageHeader_FontSize, _
                        iPageHeader_FontColor, _
                        iPageHeader_FontType, _
                        iPageHeader_BGColor, _
                        iShowFooter, _
                        iFooter_BGColor, _
                        iFooter_FontType, _
                        iFooter_FontColor, _
                        iFooter_FontColorHover, _
                        iTotalCLRows, _
                        iShowRSS, _
                        iURL_Twitter, _
                        iURL_Facebook, _
                        iURL_Myspace, _
                        iURL_Blogger)

 lcl_clID = ""

 if iCommunityLinkID <> "" then
    sCommunityLinkID = CLng(iCommunityLinkID)
 else
    sCommunityLinkID = 0
 end if

 if iIsEgovHomePage = "on" then
    sIsEgovHomePage = 1
 else
    sIsEgovHomePage = 0
 end if

'Website ----------------------------------------------------------------------
 if iWebsite_Size = "" then
  		sWebsite_Size = "'" & getCLOptionDefault("WEBSITE_SIZE") & "'"
 else
  		sWebsite_Size = "'" & dbsafe(iWebsite_Size) & "'"
 end if

 if iWebsite_Size_CustomSize = "" then
  		sWebsite_Size_CustomSize = "NULL"
 else
  		sWebsite_Size_CustomSize = "'" & dbsafe(iWebsite_Size_CustomSize) & "'"
 end if

 if iWebsite_Alignment = "" then
  		sWebsite_Alignment = "'" & getCLOptionDefault("WEBSITE_ALIGN") & "'"
 else
  		sWebsite_Alignment = "'" & dbsafe(iWebsite_Alignment) & "'"
 end if

 if iWebsite_BGColor = "" then
  		sWebsite_BGColor = "'ffffff'"
 else
  		sWebsite_BGColor = "'" & dbsafe(iWebsite_BGColor) & "'"
 end if

'Logo -------------------------------------------------------------------------
 if iShowLogo = "on" then
    sShowLogo = 1
 else
    sShowLogo = 0
 end if

 if iLogo_Filename = "" then
  		sLogo_Filename = "NULL"
 else
  		sLogo_Filename = "'" & dbsafe(iLogo_Filename) & "'"
 end if

 if iLogo_Alignment = "" then
  		sLogo_Alignment = "'" & getCLOptionDefault("WEBSITE_LOGO_ALIGN") & "'"
 else
  		sLogo_Alignment = "'" & dbsafe(iLogo_Alignment) & "'"
 end if

 if iLogo_FilenameBG = "" then
  		sLogo_FilenameBG = "NULL"
 else
  		sLogo_FilenameBG = "'" & dbsafe(iLogo_FilenameBG) & "'"
 end if

'URLs -------------------------------------------------------------------------
 if iShowRSS = "on" then
    sShowRSS = 1
 else
    sShowRSS = 0
 end if

 if iURL_Twitter = "" then
  		sURL_Twitter = "NULL"
 else
  		sURL_Twitter = "'" & dbsafe(iURL_Twitter) & "'"
 end if

 if iURL_Facebook = "" then
  		sURL_Facebook = "NULL"
 else
  		sURL_Facebook = "'" & dbsafe(iURL_Facebook) & "'"
 end if

 if iURL_Myspace = "" then
  		sURL_Myspace = "NULL"
 else
  		sURL_Myspace = "'" & dbsafe(iURL_Myspace) & "'"
 end if

 if iURL_Blogger = "" then
  		sURL_Blogger = "NULL"
 else
  		sURL_Blogger = "'" & dbsafe(iURL_Blogger) & "'"
 end if

'Top Bar ----------------------------------------------------------------------
 if iShowTopBar = "on" then
    sShowTopBar = 1
 else
    sShowTopBar = 0
 end if

 if iTopBar_BGColor = "" then
  		sTopBar_BGColor = "'ffffff'"
 else
  		sTopBar_BGColor = "'" & dbsafe(iTopBar_BGColor) & "'"
 end if

 if iTopBar_FontType = "" then
  		sTopBar_FontType = "'" & getCLOptionDefault("TOPBAR_FONTTYPE") & "'"
 else
  		sTopBar_FontType = "'" & dbsafe(iTopBar_FontType) & "'"
 end if

 if iTopBar_FontColor = "" then
  		sTopBar_FontColor = "'000000'"
 else
  		sTopBar_FontColor = "'" & dbsafe(iTopBar_FontColor) & "'"
 end if

 if iTopBar_FontColorHover = "" then
  		sTopBar_FontColorHover = "'000000'"
 else
  		sTopBar_FontColorHover = "'" & dbsafe(iTopBar_FontColorHover) & "'"
 end if

'Side Menubar -----------------------------------------------------------------
 if iShowSideMenuBar = "on" then
    sShowSideMenuBar = 1
 else
    sShowSideMenuBar = 0
 end if

 if iSideMenuBar_Alignment = "" then
  		sSideMenuBar_Alignment = "'" & getCLOptionDefault("SIDEMENUBAR_ALIGN") & "'"
 else
  		sSideMenuBar_Alignment = "'" & dbsafe(iSideMenuBar_Alignment) & "'"
 end if

 if iSideMenuOption_BGColor = "" then
  		sSideMenuOption_BGColor = "'efefef'"
 else
  		sSideMenuOption_BGColor = "'" & dbsafe(iSideMenuOption_BGColor) & "'"
 end if

 if iSideMenuOption_BGColorHover = "" then
  		sSideMenuOption_BGColorHover = "'c0c0c0'"
 else
  		sSideMenuOption_BGColorHover = "'" & dbsafe(iSideMenuOption_BGColorHover) & "'"
 end if

 if iSideMenuOption_Alignment = "" then
  		sSideMenuOption_Alignment = "'" & getCLOptionDefault("SIDEMENUOPT_TEXTALIGN") & "'"
 else
  		sSideMenuOption_Alignment = "'" & dbsafe(iSideMenuOption_Alignment) & "'"
 end if

 if iSideMenuOption_FontType = "" then
  		sSideMenuOption_FontType = "'" & getCLOptionDefault("SIDEMENUOPT_FONTTYPE") & "'"
 else
  		sSideMenuOption_FontType = "'" & dbsafe(iSideMenuOption_FontType) & "'"
 end if

 if iSideMenuOption_FontColor = "" then
  		sSideMenuOption_FontColor = "'000000'"
 else
  		sSideMenuOption_FontColor = "'" & dbsafe(iSideMenuOption_FontColor) & "'"
 end if

 if iSideMenuOption_FontColorHover = "" then
  		sSideMenuOption_FontColorHover = "'000000'"
 else
  		sSideMenuOption_FontColorHover = "'" & dbsafe(iSideMenuOption_FontColorHover) & "'"
 end if

'Page Header ------------------------------------------------------------------
 if iShowPageHeader = "on" then
    sShowPageHeader = 1
 else
    sShowPageHeader = 0
 end if

 if iPageHeader_Alignment = "" then
  		sPageHeader_Alignment = "'" & getCLOptionDefault("PAGEHEADER_ALIGN") & "'"
 else
  		sPageHeader_Alignment = "'" & dbsafe(iPageHeader_Alignment) & "'"
 end if

 if iPageHeader_FontSize = "" then
    sPageHeader_FontSize = "12"
 else
    sPageHeader_FontSize = "'" & dbsafe(iPageHeader_FontSize) & "'"
 end if

 if iPageHeader_FontColor = "" then
  		sPageHeader_FontColor = "'000000'"
 else
  		sPageHeader_FontColor = "'" & dbsafe(iPageHeader_FontColor) & "'"
 end if

 if iPageHeader_FontType = "" then
  		sPageHeader_FontType = "'" & getCLOptionDefault("PAGEHEADER_FONTTYPE") & "'"
 else
  		sPageHeader_FontType = "'" & dbsafe(iPageHeader_FontType) & "'"
 end if

 if iPageHeader_BGColor = "" then
  		sPageHeader_BGColor = "'efefef'"
 else
  		sPageHeader_BGColor = "'" & dbsafe(iPageHeader_BGColor) & "'"
 end if

'Footer -----------------------------------------------------------------------
 if iShowFooter = "on" then
    sShowFooter = 1
 else
    sShowFooter = 0
 end if

 if iFooter_BGColor = "" then
  		sFooter_BGColor = "'ffffff'"
 else
  		sFooter_BGColor = "'" & dbsafe(iFooter_BGColor) & "'"
 end if

 if iFooter_FontType = "" then
  		sFooter_FontType = "'" & getCLOptionDefault("FOOTER_FONTTYPE") & "'"
 else
  		sFooter_FontType = "'" & dbsafe(iFooter_FontType) & "'"
 end if

 if iFooter_FontColor = "" then
  		sFooter_FontColor = "'000000'"
 else
  		sFooter_FontColor = "'" & dbsafe(iFooter_FontColor) & "'"
 end if

 if iFooter_FontColorHover = "" then
  		sFooter_FontColorHover = "'000000'"
 else
  		sFooter_FontColorHover = "'" & dbsafe(iFooter_FontColorHover) & "'"
 end if

'Update the Community Link ----------------------------------------------------
 if sCommunityLinkID > 0 then
  		sSQL = "UPDATE egov_communitylink SET "
    sSQL = sSQL & "lastmodifiedbyid = "              & session("userid")                   & ", "
    sSQL = sSQL & "lastmodifiedbydate = '"           & dbsafe(ConvertDateTimetoTimeZone()) & "', "
    sSQL = sSQL & "isEgovHomePage = "                & sIsEgovHomePage                     & ", "
    sSQL = sSQL & "website_size = "                  & sWebsite_Size                       & ", "
    sSQL = sSQL & "website_size_customsize = "       & sWebsite_Size_CustomSize            & ", "
    sSQL = sSQL & "website_alignment = "             & sWebsite_Alignment                  & ", "
    sSQL = sSQL & "website_bgcolor = "               & sWebsite_BGColor                    & ", "
    sSQL = sSQL & "showlogo = "                      & sShowLogo                           & ", "
    sSQL = sSQL & "logo_filename = "                 & sLogo_Filename                      & ", "
    sSQL = sSQL & "logo_filenamebg = "               & sLogo_FilenameBG                    & ", "
    sSQL = sSQL & "logo_alignment = "                & sLogo_Alignment                     & ", "
    sSQL = sSQL & "showtopbar = "                    & sShowTopBar                         & ", "
    sSQL = sSQL & "topbar_bgcolor = "                & sTopBar_BGColor                     & ", "
    sSQL = sSQL & "topbar_fonttype = "               & sTopBar_FontType                    & ", "
    sSQL = sSQL & "topbar_fontcolor = "              & sTopBar_FontColor                   & ", "
    sSQL = sSQL & "topbar_fontcolorhover = "         & sTopBar_FontColorHover              & ", "
    sSQL = sSQL & "showsidemenubar = "               & sShowSideMenuBar                    & ", "
    sSQL = sSQL & "sidemenubar_alignment = "         & sSideMenuBar_Alignment              & ", "
    sSQL = sSQL & "sidemenuoption_bgcolor = "        & sSideMenuOption_BGColor             & ", "
    sSQL = sSQL & "sidemenuoption_bgcolorhover = "   & sSideMenuOption_BGColorHover        & ", "
    sSQL = sSQL & "sidemenuoption_alignment = "      & sSideMenuOption_Alignment           & ", "
    sSQL = sSQL & "sidemenuoption_fonttype = "       & sSideMenuOption_FontType            & ", "
    sSQL = sSQL & "sidemenuoption_fontcolor = "      & sSideMenuOption_FontColor           & ", "
    sSQL = sSQL & "sidemenuoption_fontcolorhover = " & sSideMenuOption_FontColorHover      & ", "
    sSQL = sSQL & "showpageheader = "                & sShowPageHeader                     & ", "
    sSQL = sSQL & "pageheader_alignment = "          & sPageHeader_Alignment               & ", "
    sSQL = sSQL & "pageheader_fontsize = "           & sPageHeader_FontSize                & ", "
    sSQL = sSQL & "pageheader_fontcolor = "          & sPageHeader_FontColor               & ", "
    sSQL = sSQL & "pageheader_fonttype = "           & sPageHeader_FontType                & ", "
    sSQL = sSQL & "pageheader_bgcolor = "            & sPageHeader_BGColor                 & ", "
    sSQL = sSQL & "showfooter = "                    & sShowfooter                         & ", "
    sSQL = sSQL & "footer_bgcolor = "                & sFooter_BGColor                     & ", "
    sSQL = sSQL & "footer_fonttype = "               & sFooter_FontType                    & ", "
    sSQL = sSQL & "footer_fontcolor = "              & sFooter_FontColor                   & ", "
    sSQL = sSQL & "footer_fontcolorhover = "         & sFooter_FontColorHover              & ", "
    sSQL = sSQL & "showRSS = "                       & sShowRSS                            & ", "
    sSQL = sSQL & "url_twitter = "                   & sURL_Twitter                        & ", "
    sSQL = sSQL & "url_facebook = "                  & sURL_Facebook                       & ", "
    sSQL = sSQL & "url_myspace = "                   & sURL_Myspace                        & ", "
    sSQL = sSQL & "url_blogger = "                   & sURL_Blogger
    sSQL = sSQL & " WHERE communitylinkid = " & sCommunityLinkID
'dtb_debug(sSQL)
  		set oCLUpdate = Server.CreateObject("ADODB.Recordset")
	  	oCLUpdate.Open sSQL, Application("DSN"), 3, 1

    set oCLUpdate = nothing

    lcl_success = "SU"

'------------------------------------------------------------------------------
 else  'New Community Link
'------------------------------------------------------------------------------

 		'Insert the new Community Link
  		sSQL = "INSERT INTO egov_communitylink ("
    sSQL = sSQL & "orgid, "
    sSQL = sSQL & "lastmodifiedbyid, "
    sSQL = sSQL & "lastmodifiedbydate, "
    sSQL = sSQL & "isEgovHomePage, "
    sSQL = sSQL & "website_size, "
    sSQL = sSQL & "website_size_customsize, "
    sSQL = sSQL & "website_alignment, "
    sSQL = sSQL & "website_bgcolor, "
    sSQL = sSQL & "showlogo, "
    sSQL = sSQL & "logo_filename, "
    sSQL = sSQL & "logo_filenamebg, "
    sSQL = sSQL & "logo_alignment, "
    sSQL = sSQL & "showtopbar, "
    sSQL = sSQL & "topbar_bgcolor, "
    sSQL = sSQL & "topbar_fonttype, "
    sSQL = sSQL & "topbar_fontcolor, "
    sSQL = sSQL & "topbar_fontcolorhover, "
    sSQL = sSQL & "showsidemenubar, "
    sSQL = sSQL & "sidemenubar_alignment, "
    sSQL = sSQL & "sidemenuoption_bgcolor, "
    sSQL = sSQL & "sidemenuoption_bgcolorhover, "
    sSQL = sSQL & "sidemenuoption_alignment, "
    sSQL = sSQL & "sidemenuoption_fonttype, "
    sSQL = sSQL & "sidemenuoption_fontcolor, "
    sSQL = sSQL & "sidemenuoption_fontcolorhover, "
    sSQL = sSQL & "showpageheader, "
    sSQL = sSQL & "pageheader_alignment, "
    sSQL = sSQL & "pageheader_fontsize, "
    sSQL = sSQL & "pageheader_fontcolor, "
    sSQL = sSQL & "pageheader_fonttype, "
    sSQL = sSQL & "pageheader_bgcolor, "
    sSQL = sSQL & "showfooter, "
    sSQL = sSQL & "footer_bgcolor, "
    sSQL = sSQL & "footer_fonttype, "
    sSQL = sSQL & "footer_fontcolor, "
    sSQL = sSQL & "footer_fontcolorhover, "
    sSQL = sSQL & "showRSS, "
    sSQL = sSQL & "url_twitter, "
    sSQL = sSQL & "url_facebook, "
    sSQL = sSQL & "url_myspace, "
    sSQL = sSQL & "url_blogger "
    sSQL = sSQL & ") VALUES ("
    sSQL = sSQL & session("orgid")               & ", "
    sSQL = sSQL & session("userid")              & ", "
    sSQL = sSQL & "'" & dbsafe(ConvertDateTimetoTimeZone()) & "', "
    sSQL = sSQL & IsEgovHomePage                 & ", "
    sSQL = sSQL & sWebsite_Size                  & ", "
    sSQL = sSQL & sWebsite_Size_CustomSize       & ", "
    sSQL = sSQL & sWebsite_Alignment             & ", "
    sSQL = sSQL & sWebsite_BGColor               & ", "
    sSQL = sSQL & sShowLogo                      & ", "
    sSQL = sSQL & sLogo_Filename                 & ", "
    sSQL = sSQL & sLogo_FilenameBG               & ", "
    sSQL = sSQL & sLogo_Alignment                & ", "
    sSQL = sSQL & sShowTopBar                    & ", "
    sSQL = sSQL & sTopBar_BGColor                & ", "
    sSQL = sSQL & sTopBar_FontType               & ", "
    sSQL = sSQL & sTopBar_FontColor              & ", "
    sSQL = sSQL & sTopBar_FontColorHover         & ", "
    sSQL = sSQL & sShowSideMenubar               & ", "
    sSQL = sSQL & sSideMenubar_Alignment         & ", "
    sSQL = sSQL & sSideMenuOption_BGColor        & ", "
    sSQL = sSQL & sSideMenuOption_BGColorHover   & ", "
    sSQL = sSQL & sSideMenuOption_Alignment      & ", "
    sSQL = sSQL & sSideMenuOption_FontType       & ", "
    sSQL = sSQL & sSideMenuOption_FontColor      & ", "
    sSQL = sSQL & sSideMenuOption_FontColorHover & ", "
    sSQL = sSQL & sShowPageHeader                & ", "
    sSQL = sSQL & sPageHeader_Alignment          & ", "
    sSQL = sSQL & sPageHeader_FontSize           & ", "
    sSQL = sSQL & sPageHeader_FontColor          & ", "
    sSQL = sSQL & sPageHeader_FontType           & ", "
    sSQL = sSQL & sPageHeader_BGColor            & ", "
    sSQL = sSQL & sShowFooter                    & ", "
    sSQL = sSQL & sFooter_BGColor                & ", "
    sSQL = sSQL & sFooter_FontType               & ", "
    sSQL = sSQL & sFooter_FontColor              & ", "
    sSQL = sSQL & sFooter_FontColorHover         & ", "
    sSQL = sSQL & sShowRSS                       & ", "
    sSQL = sSQL & sURL_Twitter                   & ", "
    sSQL = sSQL & sURL_Facebook                  & ", "
    sSQL = sSQL & sURL_Myspace                   & ", "
    sSQL = sSQL & sURL_Blogger
    sSQL = sSQL & ")"

    lcl_success = "SA"

    if iAction = "ADD" then
    		'Get the BlogID
   	  	lcl_communitylinkid = RunIdentityInsert(sSQL)

       lcl_clID = "&id=" & lcl_communitylinkid
    end if
 end if

'Community Link Options (cleanup) ---------------------------------------------
 deleteCLFeaturesByOrgID session("orgid")

 if iTotalCLRows > 0 then
    for i = 1 to iTotalCLRows
       lcl_isCommunityLinkOn = getRequestValue("showSection_CL_" & i,    "BIT",    "")
       lcl_isSavvyOn         = getRequestValue("showSection_SAVVY_" & i, "BIT",    "")

      'If either the Community Link OR Savvy/IFRAME options have been set to display then insert the record
       if lcl_isCommunityLinkOn OR lcl_isSavvyOn then

         'General Options -----------------------------------------------------
          if request("featurename_" & i) <> "" then
             lcl_featurename = request("featurename_" & i)
          else
             lcl_featurename = request("featurename_original_" & i)
          end if

          if request("portalcolumn_" & i) <> "" then
             lcl_portalcolumn = request("portalcolumn_" & i)
          else
             lcl_portalcolumn = 1
          end if

          if request("displayorder_" & i) <> "" then
             lcl_displayorder = request("displayorder_" & i)
          else
             lcl_displayorder = 1
          end if

          if request("rss_feedid_" & i) <> "" then
             lcl_rss_feedid = request("rss_feedid_" & i)
          else
             lcl_rss_feedid = 0
          end if

          if request("query_filter_" & i) <> "" then
             lcl_query_filter = request("query_filter_" & i)
          else
             lcl_query_filter = ""
          end if

dtb_debug("before: lcl_viewall_urltype_CL: [" & request("sectionlinks_viewall_urltype_CL_" & i) & "] - lcl_viewall_url_CL: [" & request("sectionlinks_viewall_url_CL_" & i) & "] - lcl_viewall_url_wintype_CL: [" & request("sectionlinks_viewall_url_wintype_CL_" & i) & "]")

         'Community Link Options (setup) -----------------------------------------
          lcl_showSectionBorder_CL           = getRequestValue("showsectionborder_CL_"                & i, "BIT",                 "")
          lcl_sectionBorderColor_CL          = getRequestValue("sectionbordercolor_CL_"               & i, "BGCOLOR",             "000000")
          lcl_sectionBackgroundColor_CL      = getRequestValue("sectionbackgroundcolor_CL_"           & i, "",                    "")
          lcl_sectionHeader_BGColor_CL       = getRequestValue("sectionheader_bgcolor_CL_"            & i, "BGCOLOR",             "")
          lcl_sectionHeader_LineColor_CL     = getRequestValue("sectionheader_linecolor_CL_"          & i, "BGCOLOR",             "000000")
          lcl_sectionHeader_FontType_CL      = getRequestValue("sectionheader_fonttype_CL_"           & i, "FONTTYPE",            getCLOptionDefault("SECTIONHEADER_FONTTYPE"))
          lcl_sectionHeader_FontColor_CL     = getRequestValue("sectionheader_fontcolor_CL_"          & i, "FONTCOLOR",           "")
          lcl_sectionHeader_FontSize_CL      = getRequestValue("sectionheader_fontsize_CL_"           & i, "FONTSIZE",            "11")
          lcl_sectionHeader_FontSize_CL      = getRequestValue("sectionheader_fontsize_CL_"           & i, "FONTSIZE",            "11")
          lcl_sectionHeader_isBold_CL        = getRequestValue("sectionheader_isbold_CL_"             & i, "BIT",                 "")
          lcl_sectionHeader_isItalic_CL      = getRequestValue("sectionheader_isitalic_CL_"           & i, "BIT",                 "")
          lcl_sectionText_BGColor_CL         = getRequestValue("sectiontext_bgcolor_CL_"              & i, "BGCOLOR",             "")
          lcl_sectionText_BGColorHover_CL    = getRequestValue("sectiontext_bgcolorhover_CL_"         & i, "BGCOLOR",             "")
          lcl_sectionText_FontType_CL        = getRequestValue("sectiontext_fonttype_CL_"             & i, "FONTTYPE",            getCLOptionDefault("SECTIONTEXT_FONTTYPE"))
          lcl_sectionText_FontColor_CL       = getRequestValue("sectiontext_fontcolor_CL_"            & i, "FONTCOLOR",           "")
          lcl_sectionText_FontColorHover_CL  = getRequestValue("sectiontext_fontcolorhover_CL_"       & i, "FONTCOLOR",           "")
          lcl_sectionText_FontSize_CL        = getRequestValue("sectiontext_fontsize_CL_"             & i, "FONTSIZE",            "11")
          lcl_numListItemsShown_CL           = getRequestValue("numListItemsShown_CL_"                & i, "LISTITEMS",           request("numListItemsShown_original_" & i))
          lcl_sectionLinks_Alignment_CL      = getRequestValue("sectionlinks_alignment_CL_"           & i, "ALIGNMENT",           getCLOptionDefault("SECTIONLINKS_ALIGN"))
          lcl_sectionLinks_FontType_CL       = getRequestValue("sectionlinks_fonttype_CL_"            & i, "FONTTYPE",            getCLOptionDefault("SECTIONLINKS_FONTTYPE"))
          lcl_sectionLinks_FontColor_CL      = getRequestValue("sectionlinks_fontcolor_CL_"           & i, "BGCOLOR",             "800000")
          lcl_sectionLinks_FontColorHover_CL = getRequestValue("sectionlinks_fontcolorhover_CL_"      & i, "BGCOLOR",             "800000")
          lcl_viewall_urltype_CL             = getRequestValue("sectionlinks_viewall_urltype_CL_"     & i, "VIEWALL_URLTYPE",     getCLOptionDefault("VIEWALL_URLTYPE"))
          lcl_viewall_url_CL                 = getRequestValue("sectionlinks_viewall_url_CL_"         & i, "",                    "")
          lcl_viewall_url_wintype_CL         = getRequestValue("sectionlinks_viewall_url_wintype_CL_" & i, "VIEWALL_URL_WINTYPE", getCLOptionDefault("VIEWALL_URL_WINTYPE"))

         'Savvy/IFRAME Options (setup) -------------------------------------------
          lcl_showSectionBorder_SAVVY           = getRequestValue("showsectionborder_SAVVY_"                & i, "BIT",                 "")
          lcl_sectionBorderColor_SAVVY          = getRequestValue("sectionbordercolor_SAVVY_"               & i, "BGCOLOR",             "000000")
          lcl_sectionBackgroundColor_SAVVY      = getRequestValue("sectionbackgroundcolor_SAVVY_"           & i, "",                    "")
          lcl_sectionHeader_BGColor_SAVVY       = getRequestValue("sectionheader_bgcolor_SAVVY_"            & i, "BGCOLOR",             "")
          lcl_sectionHeader_LineColor_SAVVY     = getRequestValue("sectionheader_linecolor_SAVVY_"          & i, "BGCOLOR",             "000000")
          lcl_sectionHeader_FontType_SAVVY      = getRequestValue("sectionheader_fonttype_SAVVY_"           & i, "FONTTYPE",            getCLOptionDefault("SECTIONHEADER_FONTTYPE"))
          lcl_sectionHeader_FontColor_SAVVY     = getRequestValue("sectionheader_fontcolor_SAVVY_"          & i, "FONTCOLOR",           "")
          lcl_sectionHeader_FontSize_SAVVY      = getRequestValue("sectionheader_fontsize_SAVVY_"           & i, "FONTSIZE",            "11")
          lcl_sectionHeader_isBold_SAVVY        = getRequestValue("sectionheader_isbold_SAVVY_"             & i, "BIT",                 "")
          lcl_sectionHeader_isItalic_SAVVY      = getRequestValue("sectionheader_isitalic_SAVVY_"           & i, "BIT",                 "")
          lcl_sectionText_BGColor_SAVVY         = getRequestValue("sectiontext_bgcolor_SAVVY_"              & i, "BGCOLOR",             "")
          lcl_sectionText_BGColorHover_SAVVY    = getRequestValue("sectiontext_bgcolorhover_SAVVY_"         & i, "BGCOLOR",             "")
          lcl_sectionText_FontType_SAVVY        = getRequestValue("sectiontext_fonttype_SAVVY_"             & i, "FONTTYPE",            getCLOptionDefault("SECTIONTEXT_FONTTYPE"))
          lcl_sectionText_FontColor_SAVVY       = getRequestValue("sectiontext_fontcolor_SAVVY_"            & i, "FONTCOLOR",           "")
          lcl_sectionText_FontColorHover_SAVVY  = getRequestValue("sectiontext_fontcolorhover_SAVVY_"       & i, "FONTCOLOR",           "")
          lcl_sectionText_FontSize_SAVVY        = getRequestValue("sectiontext_fontsize_SAVVY_"             & i, "FONTSIZE",            "11")
          lcl_numListItemsShown_SAVVY           = getRequestValue("numListItemsShown_SAVVY_"                & i, "LISTITEMS",           request("numListItemsShown_original_" & i))
          lcl_sectionLinks_Alignment_SAVVY      = getRequestValue("sectionlinks_alignment_SAVVY_"           & i, "ALIGNMENT",           getCLOptionDefault("SECTIONLINKS_ALIGN"))
          lcl_sectionLinks_FontType_SAVVY       = getRequestValue("sectionlinks_fonttype_SAVVY_"            & i, "FONTTYPE",            getCLOptionDefault("SECTIONLINKS_FONTTYPE"))
          lcl_sectionLinks_FontColor_SAVVY      = getRequestValue("sectionlinks_fontcolor_SAVVY_"           & i, "BGCOLOR",             "800000")
          lcl_sectionLinks_FontColorHover_SAVVY = getRequestValue("sectionlinks_fontcolorhover_SAVVY_"      & i, "BGCOLOR",             "800000")
          lcl_viewall_urltype_SAVVY             = getRequestValue("sectionlinks_viewall_urltype_SAVVY_"     & i, "VIEWALL_URLTYPE",     getCLOptionDefault("VIEWALL_URLTYPE"))
          lcl_viewall_url_SAVVY                 = getRequestValue("sectionlinks_viewall_url_SAVVY_"         & i, "",                    "")
          lcl_viewall_url_wintype_SAVVY         = getRequestValue("sectionlinks_viewall_url_wintype_SAVVY_" & i, "VIEWALL_URL_WINTYPE", getCLOptionDefault("VIEWALL_URL_WINTYPE"))
dtb_debug("lcl_viewall_urltype_CL: [" & lcl_viewall_urltype_CL & "] - lcl_viewall_url_CL: [" & lcl_viewall_url_CL & "] - lcl_viewall_url_wintype_CL: [" & lcl_viewall_url_wintype_CL & "]")
          insertCLFeatures session("orgid"), _
                           request("featureid_" & i), _
                           lcl_featurename, _
                           lcl_portalcolumn, _
                           lcl_displayorder, _
                           lcl_rss_feedid, _
                           lcl_numListItemsShown_CL, _
                           lcl_numListItemsShown_SAVVY, _
                           lcl_isCommunityLinkOn, _
                           lcl_isSavvyOn, _
                           lcl_showSectionBorder_CL, _
                           lcl_showSectionBorder_SAVVY, _
                           lcl_sectionBorderColor_CL, _
                           lcl_sectionBorderColor_SAVVY, _
                           lcl_sectionBackgroundColor_CL, _
                           lcl_sectionBackgroundColor_SAVVY, _
                           lcl_sectionHeader_BGColor_CL, _
                           lcl_sectionHeader_BGColor_SAVVY, _
                           lcl_sectionHeader_LineColor_CL, _
                           lcl_sectionHeader_LineColor_SAVVY, _
                           lcl_sectionHeader_FontType_CL, _
                           lcl_sectionHeader_FontType_SAVVY, _
                           lcl_sectionHeader_FontColor_CL, _
                           lcl_sectionHeader_FontColor_SAVVY, _
                           lcl_sectionHeader_FontSize_CL, _
                           lcl_sectionHeader_FontSize_SAVVY, _
                           lcl_sectionHeader_isBold_CL, _
                           lcl_sectionHeader_isBold_SAVVY, _
                           lcl_sectionHeader_isItalic_CL, _
                           lcl_sectionHeader_isItalic_SAVVY, _
                           lcl_sectionText_BGColor_CL, _
                           lcl_sectionText_BGColor_SAVVY, _
                           lcl_sectionText_BGColorHover_CL, _
                           lcl_sectionText_BGColorHover_SAVVY, _
                           lcl_sectionText_FontType_CL, _
                           lcl_sectionText_FontType_SAVVY, _
                           lcl_sectionText_FontColor_CL, _
                           lcl_sectionText_FontColor_SAVVY, _
                           lcl_sectionText_FontColorHover_CL, _
                           lcl_sectionText_FontColorHover_SAVVY, _
                           lcl_sectionText_FontSize_CL, _
                           lcl_sectionText_FontSize_SAVVY, _
                           lcl_sectionLinks_Alignment_CL, _
                           lcl_sectionLinks_Alignment_SAVVY, _
                           lcl_sectionLinks_FontType_CL, _
                           lcl_sectionLinks_FontType_SAVVY, _
                           lcl_sectionLinks_FontColor_CL, _
                           lcl_sectionLinks_FontColor_SAVVY, _
                           lcl_sectionLinks_FontColorHover_CL, _
                           lcl_sectionLinks_FontColorHover_SAVVY, _
                           lcl_viewall_urltype_CL, _
                           lcl_viewall_urltype_SAVVY, _
                           lcl_viewall_url_CL, _
                           lcl_viewall_url_SAVVY, _
                           lcl_viewall_url_wintype_CL, _
                           lcl_viewall_url_wintype_SAVVY, _
                           lcl_query_filter

       end if
    next
 end if

 response.redirect "communitylink_maint.asp?success=" & lcl_success & lcl_clID

end sub

'------------------------------------------------------------------------------
sub deleteCLFeaturesByOrgID(p_orgid)

  if p_orgid <> "" then
     sSQL = "DELETE FROM egov_communitylink_displayorgfeatures WHERE orgid = " & p_orgid

   		set oDeleteCLFeatures = Server.CreateObject("ADODB.Recordset")
    	oDeleteCLFeatures.Open sSQL, Application("DSN"), 3, 1

     set oDeleteCLFeatures = nothing

  end if

end sub

'------------------------------------------------------------------------------
sub insertCLFeatures(p_orgid, _
                     p_featureid, _
                     p_featurename, _
                     p_portalcolumn, _
                     p_displayorder, _
                     p_rss_feedid, _
                     p_numListItemsShown_CL, _
                     p_numListItemsShown_SAVVY, _
                     p_isCommunityLinkOn, _
                     p_isSavvyOn, _
                     p_showSectionBorder_CL, _
                     p_showSectionBorder_SAVVY, _
                     p_sectionBorderColor_CL, _
                     p_sectionBorderColor_SAVVY, _
                     p_sectionBackgroundColor_CL, _
                     p_sectionBackgroundColor_SAVVY, _
                     p_sectionHeader_BGColor_CL, _
                     p_sectionHeader_BGColor_SAVVY, _
                     p_sectionHeader_LineColor_CL, _
                     p_sectionHeader_LineColor_SAVVY, _
                     p_sectionHeader_FontType_CL, _
                     p_sectionHeader_FontType_SAVVY, _
                     p_sectionHeader_FontColor_CL, _
                     p_sectionHeader_FontColor_SAVVY, _
                     p_sectionHeader_FontSize_CL, _
                     p_sectionHeader_FontSize_SAVVY, _
                     p_sectionHeader_isBold_CL, _
                     p_sectionHeader_isBold_SAVVY, _
                     p_sectionHeader_isItalic_CL, _
                     p_sectionHeader_isItalic_SAVVY, _
                     p_sectionText_BGColor_CL, _
                     p_sectionText_BGColor_SAVVY, _
                     p_sectionText_BGColorHover_CL, _
                     p_sectionText_BGColorHover_SAVVY, _
                     p_sectionText_FontType_CL, _
                     p_sectionText_FontType_SAVVY, _
                     p_sectionText_FontColor_CL, _
                     p_sectionText_FontColor_SAVVY, _
                     p_sectionText_FontColorHover_CL, _
                     p_sectionText_FontColorHover_SAVVY, _
                     p_sectionText_FontSize_CL, _
                     p_sectionText_FontSize_SAVVY, _
                     p_sectionLinks_Alignment_CL, _
                     p_sectionLinks_Alignment_SAVVY, _
                     p_sectionLinks_FontType_CL, _
                     p_sectionLinks_FontType_SAVVY, _
                     p_sectionLinks_FontColor_CL, _
                     p_sectionLinks_FontColor_SAVVY, _
                     p_sectionLinks_FontColorHover_CL, _
                     p_sectionLinks_FontColorHover_SAVVY, _
                     p_viewall_urltype_CL, _
                     p_viewall_urltype_SAVVY, _
                     p_viewall_url_CL, _
                     p_viewall_url_SAVVY, _
                     p_viewall_url_wintype_CL, _
                     p_viewall_url_wintype_SAVVY, _
                     p_query_filter)

  sSQL = "INSERT INTO egov_communitylink_displayorgfeatures ("
  sSQL = sSQL & "orgid, "
  sSQL = sSQL & "featureid, "
  sSQL = sSQL & "featurename, "
  sSQL = sSQL & "portalcolumn, "
  sSQL = sSQL & "displayorder, "
  sSQL = sSQL & "rss_feedid, "
  sSQL = sSQL & "numListItemsShown_CL, "
  sSQL = sSQL & "numListItemsShown_SAVVY, "
  sSQL = sSQL & "isCommunityLinkOn, "
  sSQL = sSQL & "isSavvyOn, "
  sSQL = sSQL & "showsectionborder_CL, "
  sSQL = sSQL & "showsectionborder_SAVVY, "
  sSQL = sSQL & "sectionbordercolor_CL, "
  sSQL = sSQL & "sectionbordercolor_SAVVY, "
  sSQL = sSQL & "sectionbackgroundcolor_CL, "
  sSQL = sSQL & "sectionbackgroundcolor_SAVVY, "
  sSQL = sSQL & "sectionheader_bgcolor_CL, "
  sSQL = sSQL & "sectionheader_bgcolor_SAVVY, "
  sSQL = sSQL & "sectionheader_linecolor_CL, "
  sSQL = sSQL & "sectionheader_linecolor_SAVVY, "
  sSQL = sSQL & "sectionheader_fonttype_CL, "
  sSQL = sSQL & "sectionheader_fonttype_SAVVY, "
  sSQL = sSQL & "sectionheader_fontcolor_CL, "
  sSQL = sSQL & "sectionheader_fontcolor_SAVVY, "
  sSQL = sSQL & "sectionheader_fontsize_CL, "
  sSQL = sSQL & "sectionheader_fontsize_SAVVY, "
  sSQL = sSQL & "sectionheader_isbold_CL, "
  sSQL = sSQL & "sectionheader_isbold_SAVVY, "
  sSQL = sSQL & "sectionheader_isitalic_CL, "
  sSQL = sSQL & "sectionheader_isitalic_SAVVY, "
  sSQL = sSQL & "sectiontext_bgcolor_CL, "
  sSQL = sSQL & "sectiontext_bgcolor_SAVVY, "
  sSQL = sSQL & "sectiontext_bgcolorhover_CL, "
  sSQL = sSQL & "sectiontext_bgcolorhover_SAVVY, "
  sSQL = sSQL & "sectiontext_fonttype_CL, "
  sSQL = sSQL & "sectiontext_fonttype_SAVVY, "
  sSQL = sSQL & "sectiontext_fontcolor_CL, "
  sSQL = sSQL & "sectiontext_fontcolor_SAVVY, "
  sSQL = sSQL & "sectiontext_fontcolorhover_CL, "
  sSQL = sSQL & "sectiontext_fontcolorhover_SAVVY, "
  sSQL = sSQL & "sectiontext_fontsize_CL, "
  sSQL = sSQL & "sectiontext_fontsize_SAVVY, "
  sSQL = sSQL & "sectionlinks_alignment_CL, "
  sSQL = sSQL & "sectionlinks_alignment_SAVVY, "
  sSQL = sSQL & "sectionlinks_fonttype_CL, "
  sSQL = sSQL & "sectionlinks_fonttype_SAVVY, "
  sSQL = sSQL & "sectionlinks_fontcolor_CL, "
  sSQL = sSQL & "sectionlinks_fontcolor_SAVVY, "
  sSQL = sSQL & "sectionlinks_fontcolorhover_CL, "
  sSQL = sSQL & "sectionlinks_fontcolorhover_SAVVY, "
  sSQL = sSQL & "viewall_urltype_CL, "
  sSQL = sSQL & "viewall_urltype_SAVVY, "
  sSQL = sSQL & "viewall_url_CL, "
  sSQL = sSQL & "viewall_url_SAVVY, "
  sSQL = sSQL & "viewall_url_wintype_CL, "
  sSQL = sSQL & "viewall_url_wintype_SAVVY, "
  sSQL = sSQL & "query_filter "
  sSQL = sSQL & ") VALUES ("
  sSQL = sSQL &       p_orgid                                     & ", "
  sSQL = sSQL &       p_featureid                                 & ", "
  sSQL = sSQL & "'" & dbsafe(p_featurename)                       & "', "
  sSQL = sSQL &       p_portalcolumn                              & ", "
  sSQL = sSQL &       p_displayorder                              & ", "
  sSQL = sSQL &       p_rss_feedid                                & ", "
  sSQL = sSQL &       p_numListItemsShown_CL                      & ", "
  sSQL = sSQL &       p_numListItemsShown_SAVVY                   & ", "
  sSQL = sSQL &       p_isCommunityLinkOn                         & ", "
  sSQL = sSQL &       p_isSavvyOn                                 & ", "
  sSQL = sSQL &       p_showSectionBorder_CL                      & ", "
  sSQL = sSQL &       p_showSectionBorder_SAVVY                   & ", "
  sSQL = sSQL & "'" & dbsafe(p_sectionBorderColor_CL)             & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionBorderColor_SAVVY)          & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionBackgroundColor_CL)         & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionBackgroundColor_SAVVY)      & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_BGColor_CL)          & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_BGColor_SAVVY)       & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_LineColor_CL)        & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_LineColor_SAVVY)     & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_FontType_CL)         & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_FontType_SAVVY)      & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_FontColor_CL)        & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_FontColor_SAVVY)     & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_FontSize_CL)         & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_FontSize_SAVVY)      & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_isBold_CL)           & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_isBold_SAVVY)        & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_isItalic_CL)         & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionHeader_isItalic_SAVVY)      & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_BGColor_CL)            & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_BGColor_SAVVY)         & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_BGColorHover_CL)       & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_BGColorHover_SAVVY)    & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_FontType_CL)           & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_FontType_SAVVY)        & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_FontColor_CL)          & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_FontColor_SAVVY)       & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_FontColorHover_CL)     & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_FontColorHover_SAVVY)  & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_FontSize_CL)           & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionText_FontSize_SAVVY)        & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionLinks_Alignment_CL)         & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionLinks_Alignment_SAVVY)      & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionLinks_FontType_CL)          & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionLinks_FontType_SAVVY)       & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionLinks_FontColor_CL)         & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionLinks_FontColor_SAVVY)      & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionLinks_FontColorHover_CL)    & "', "
  sSQL = sSQL & "'" & dbsafe(p_sectionLinks_FontColorHover_SAVVY) & "', "
  sSQL = sSQL & "'" & dbsafe(p_viewall_urltype_CL)                & "', "
  sSQL = sSQL & "'" & dbsafe(p_viewall_urltype_SAVVY)             & "', "
  sSQL = sSQL & "'" & dbsafe(p_viewall_url_CL)                    & "', "
  sSQL = sSQL & "'" & dbsafe(p_viewall_url_SAVVY)                 & "', "
  sSQL = sSQL & "'" & dbsafe(p_viewall_url_wintype_CL)            & "', "
  sSQL = sSQL & "'" & dbsafe(p_viewall_url_wintype_SAVVY)         & "', "
  sSQL = sSQL & "'" & dbsafe(p_query_filter)                      & "' "
  sSQL = sSQL & ") "
dtb_debug(sSQL)
  set oInsertCLFeatures = Server.CreateObject("ADODB.Recordset")
  oInsertCLFeatures.Open sSQL, Application("DSN"), 3, 1

  set oInsertCLFeatures = nothing

end sub

'------------------------------------------------------------------------------
function getRequestValue(iFieldName, iFieldType, iDefaultValue)

  lcl_return = ""

 'Retrieve the field if a field name has been passed in and a field/value exists to pull from.
 'Otherwise, use the default value passed in.
 'If no default value has been passed in then set the default value for the field type passed in.
  if request(iFieldName) <> "" then
     if iFieldType = "BIT" then
        lcl_return = 1
     else
        lcl_return = request(iFieldName)
     end if
  else
     if iDefaultValue <> "" then
        lcl_return = iDefaultValue
     end if
  end if

  if lcl_return = "" then
    'BIT ----------------------------------------------------------------------
     if iFieldType = "BIT" then
        if iDefaultValue = "" then
           lcl_return = 0
        end if

    'BGCOLOR ------------------------------------------------------------------
     elseif iFieldType = "BGCOLOR" then
        if iDefaultValue = "" then
           lcl_return = "ffffff"
        end if

    'FONTTYPE -----------------------------------------------------------------
     elseif iFieldType = "FONTTYPE" then
        if iDefaultValue = "" then
           lcl_return = "Verdana"
        end if

    'FONTCOLOR ----------------------------------------------------------------
     elseif iFieldType = "FONTCOLOR" then
        if iDefaultValue = "" then
           lcl_return = "000000"
        end if

    'FONTSIZE ----------------------------------------------------------------
     elseif iFieldType = "FONTSIZE" then
        if iDefaultValue = "" then
           lcl_return = "11"
        end if

    'LISTITEMS ----------------------------------------------------------------
     elseif iFieldType = "LISTITEMS" then
        if iDefaultValue = "" then
           lcl_return = "0"
        end if

    'VARCHAR ------------------------------------------------------------------
     elseif iFieldType = "VARCHAR" then
        if iDefaultValue = "" then
           lcl_return = ""
        end if

    'ALIGNMENT ----------------------------------------------------------------
     elseif iFieldType = "ALIGNMENT" then
        if iDefaultValue = "" then
           lcl_return = ""
        end if

     end if
  end if

  getRequestValue = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQL = "INSERT INTO my_table_dtb(notes) VALUES('" & replace(p_value,"'","''") & "')"
		set oDTB = Server.CreateObject("ADODB.Recordset")
 	oDTB.Open sSQL, Application("DSN"), 3, 1

  set oDTB = nothing

end sub
%>