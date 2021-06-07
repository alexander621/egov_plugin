<%@ Page Language="C#" AutoEventWireup="true" %>

<%@ Import Namespace="System.Configuration" %>

<%
/// <summary>
///
/// jQuery File Tree ASP Connector
///
/// Version 1.0
///
/// Copyright (c)2008 Andrew Sweeny
/// asweeny@fit.edu
/// 24 March 2008
///
/// </summary>
	
	string dir;
	string OrgUrl;
	string UserId;
	Boolean HasInternalSecurity = false;
	string FolderClass = "";
	Int32 FolderFileCounter = 0;
    string documentRoot;

	documentRoot = ConfigurationManager.AppSettings["DocumentsRootDirectory"];
	//documentRoot = "egovlink300_docs";
	
	if(Request.Form["dir"] == null || Request.Form["dir"].Length <= 0)
		dir = "\\" + documentRoot + "\\custom\\pub\\" + common.GetVirtualDirectyName( HttpContext.Current.Request.ServerVariables["URL"].ToLower( ) ) + "\\";
	else
		dir = Request.Form["dir"];

	if ( Request["userid"] == null || Request["userid"].Length <= 0 )
		UserId = "0";
	else
		UserId = Int32.Parse( Request["userid"]).ToString(); 
	

	//Response.Write(dir + "<br />");
	dir = dir.Replace( "%20", " " );

    OrgUrl = "/" + documentRoot + "/custom/pub/" + common.GetVirtualDirectyName( HttpContext.Current.Request.ServerVariables["URL"].ToLower( ) ) + "/";

	System.IO.DirectoryInfo di = new System.IO.DirectoryInfo("e:" + dir);
	
	Response.Write("\n<ul class=\"jqueryFileTree\" style=\"display: none;\">");

	foreach ( System.IO.DirectoryInfo di_child in di.GetDirectories( ) )
	{
		if ( dir + di_child.Name != OrgUrl + "attachments" && dir + di_child.Name != OrgUrl + "pdf_forms" && dir + di_child.Name != OrgUrl + "postings_bids" )
		{
			FolderClass = "";

			if ( dir + di_child.Name == OrgUrl + "published_documents" || dir + di_child.Name == OrgUrl + "unpublished_documents" )
				FolderClass = " root";

            HasInternalSecurity = common.FolderHasInternalSecurity( dir.Replace( documentRoot, "public_documents300" ) + di_child.Name );
			if ( HasInternalSecurity )
				FolderClass = " lockedfolder";
			
			// need to check internal security here to see if they are allowed to see this folder
            if (!HasInternalSecurity || common.UserCanViewSecureFolder( UserId, dir.Replace( documentRoot, "public_documents300" ) + di_child.Name ))
			{
				Response.Write( "\n\t<li class=\"directory collapsed" + FolderClass + "\"><a class=\"folder\" href=\"#\" rel=\"" + dir + di_child.Name + "/\">" );

				// if restricted on public side put &reg; here
                if (common.FolderHasRestrictedPublicAccess( dir.Replace( documentRoot, "public_documents300" ) + di_child.Name ))
					Response.Write( "<b>&reg;</b> " );
				
				Response.Write(  di_child.Name + "</a></li>" );
				FolderFileCounter += 1;
			}
			
		}
	}
	Int32 FileCounter = 0;
	
	foreach (System.IO.FileInfo fi in di.GetFiles())
	{
		string ext = "";
		
		if(fi.Extension.Length > 1)
			ext = fi.Extension.Substring(1).ToLower();

        Response.Write( "\t<li class=\"file ext_" + ext + "\"><a href=\"" + dir.Replace( documentRoot + "/custom/pub", "public_documents300" ) + fi.Name + "\" rel=\"" + dir + fi.Name + "\" target=\"_blank\">" + fi.Name + "</a></li>\n" );
		//Response.Write( "\t<li class=\"file ext_" + ext + "\"><a href=\"" + dir.Replace("egovlink300_docs/custom/pub", "public_documents300") + fi.Name + "\" rel=\"" + dir + fi.Name + "\" target=\"_blank\">" + fi.Name + "</a></li>\n" );

		FolderFileCounter += 1;
	}
	if ( FolderFileCounter == 0 )
		Response.Write( "\n\t<li>(<i>Empty</i>)</li>" );
	
	Response.Write("</ul>");
	Response.Write("<script>showMenu(); </script>");

		
 %>