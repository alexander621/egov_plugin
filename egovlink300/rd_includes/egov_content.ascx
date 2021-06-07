<%@ Control Language="C#" AutoEventWireup="true" CodeFile="egov_content.ascx.cs" Inherits="rd_includes_egovcontent" %>

  <%

      string sOrgID = common.getOrgId();
      
      Int32 iRootCategoryID  = getFirstCategory();
      Int32 iCategoryID      = 0;
      Boolean sViewPick      = false;
      Boolean sShowViewPicks = true;

      if (Request["categoryid"] != null)
      {
          iCategoryID = Convert.ToInt32(Request["categoryid"]);
      }
      else
      {
          iCategoryID = iRootCategoryID;
      }

      if (Request.QueryString["viewpick"] != null)
      {
          try
          {
              if (Convert.ToInt32(Request.QueryString["viewpick"]) == 1)
              {
                  sViewPick = true;
              }
          }
          catch
          {
              sViewPick = false;
          }
      }

      if (iCategoryID == iRootCategoryID)
      {
          sShowViewPicks = false;
      }
      
      Response.Write("<div id=\"content\">");

                        showMemberWarning();
                        displaySubCategoryMenu(iRootCategoryID, sShowViewPicks, sViewPick, iCategoryID);

      Response.Write("</div>");
  %>
  
    <p>
    Display screen as Logged In: 
    <input type="radio" name="isLoggedin" id="isLoggedIn_yes" value="Y"<%=lcl_checked_isLoggedInYes%> />Yes
    <input type="radio" name="isLoggedIn" id="isLoggedIn_no" value="N"<%=lcl_checked_isLoggedInNo%> />No
    </p>
