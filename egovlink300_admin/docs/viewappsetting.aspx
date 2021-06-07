<%@ Page Language="C#" AutoEventWireup="true" %>

<%@ Import Namespace="System.Configuration" %>

<%

    string documentRoot;

	documentRoot = ConfigurationManager.AppSettings["DocumentsRootDirectory"];

    Response.Write( "documentRoot: " + documentRoot );
		
 %>