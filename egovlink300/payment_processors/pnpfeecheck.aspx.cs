using System;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using System.Net;
using System.IO;
using System.Text;
using System.Configuration;
using System.Security.Authentication;

/// <summary>
///--------------------------------------------------------------------------------------------------
///
///--------------------------------------------------------------------------------------------------
/// FILENAME: pnpfeecheck.aspx.cs
/// AUTHOR: Steve Loar
/// CREATED: 06/30/2010
/// COPYRIGHT: Copyright 2010 eclink, inc.
///			 All Rights Reserved.
///
/// DESCTIPTION: This gets the PNP Fee for the payment form, is called via AJAX. Needs it's parent
///				page of pnpfeecheck.aspx to work
///
/// MODIFICATION HISTORY
/// 1.0   06/30/2010	Steve Loar - INITIAL VERSION
///
///--------------------------------------------------------------------------------------------------
///
///--------------------------------------------------------------------------------------------------
/// </summary>


public partial class pnpfeecheck : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
		Response.Clear( );

		string sOrgId = common.getOrgId( );

		//log the fee check start
		string iPaymentControlNumber = common.CreatePaymentControlRow( sOrgId, "PNP Fee Check", "Fee Check Started" );

		// Check that the amount is a double
		double dAmount;

		if ( double.TryParse( Request["chkamount"], out dAmount ) == false )
		{
			// log the bad amount
			common.AddToPaymentLog( sOrgId, "PNP Fee Check", iPaymentControlNumber, "Amount passed was not currency: " + Request["chkamount"] );

			// write that the amount was bad
			Response.Write( "results=0&status=failed&errors=invalid parameter value" );
		}
		else
		{
			// Log the values passed in
			common.AddToPaymentLog( sOrgId, "PNP Fee Check", iPaymentControlNumber, "Amount Passed = " + Request["chkamount"] );

			// All is OK so try to get the fee from PNP
			string sReturnString = GetFee( sOrgId, Request["chkamount"] );

			// log the string to be returned
			common.AddToPaymentLog( sOrgId, "PNP Fee Check", iPaymentControlNumber, sReturnString );

			// write out the results so the calling script can get it
			Response.Write( sReturnString );
		}

		// Log that the script is done
		common.AddToPaymentLog( sOrgId, "PNP Fee Check", iPaymentControlNumber, "Fee Check Finished" );

		Response.End( );
    }


	//--------------------------------------------------------------------------------------------------
	private string GetFee( string iOrgId, string sAmount )
	{
		string requestStr;

		if ( CreateFeeXML( iOrgId, sAmount, out requestStr ) )
		{
			string showSendStr = requestStr;
			string sPNPURL;
			string Results = "";

			// fetch the PNP url here 
			if ( GetPNPURL( iOrgId, out sPNPURL ) )
				Results = PostData( sPNPURL, requestStr );
			else
			{
				return "results=0&status=failed&errors=Failed to get PNP URL" ; 
			}

			XmlDocument doc = new XmlDocument( );
			doc.LoadXml( Results );
			XmlElement root = doc.DocumentElement;
			XmlNodeList oNodes = root.FirstChild.FirstChild.ChildNodes;

			string sResults = "";
			sResults = "result=" + oNodes.Count.ToString( );
			foreach ( XmlNode node in oNodes )
			{
				if ( node.Name.ToLower( ) == "errors" )
				{
					if ( node.InnerText.Trim( ) != "" )
					{
						sResults += "&" + node.Name.ToLower( ) + "=" + node.InnerText;
					}
					else
						sResults += "&errors=none";
				}
				else
					sResults += "&" + node.Name.ToLower( ) + "=" + node.InnerText;
			}
			
			return sResults ;
		}
		else
		{
			return "results=0&status=failed&errors=Failed to create XML request string" ;
		}
	}


	//--------------------------------------------------------------------------------------------------
	private Boolean CreateFeeXML( string iOrgId, string sAmount, out string sXmlReturn )
	{
		XmlDocument doc = new XmlDocument( );
		string sPartnerCode;
		string sUserName;
		string sPassWord;
		string sProductId;
		string sOfficeCd;

		if ( GetPNPOptions( iOrgId, out sPartnerCode, out sUserName, out sPassWord, out sProductId, out sOfficeCd ) )
		{
			XmlElement elemRoot = doc.CreateElement( "root" );
			XmlElement elemData = doc.CreateElement( "data" );
			XmlElement elemStruct = doc.CreateElement( "struct" );
			doc.AppendChild( elemRoot );
			elemRoot.AppendChild( elemData );
			elemData.AppendChild( elemStruct );

			//Insert data here
			InsertDataNode( doc, elemStruct, "PartnerCD", sPartnerCode );
			InsertDataNode( doc, elemStruct, "UserName", sUserName );
			InsertDataNode( doc, elemStruct, "Password", sPassWord );
			InsertDataNode( doc, elemStruct, "Action", "Fee" );
			InsertDataNode( doc, elemStruct, "ProductID", sProductId );
			InsertDataNode( doc, elemStruct, "Amount", sAmount );
			InsertDataNode( doc, elemStruct, "PaymentDevice", "CreditCard" );

			sXmlReturn = doc.OuterXml;
			return true;
		}
		else
		{
			//Response.Write( "No PNP Options " );
			sXmlReturn = "failed";
			return false;
		}
	}


	//--------------------------------------------------------------------------------------------------
	private void InsertDataNode( XmlDocument doc, XmlElement parentElem, string nodeName, string nodeValue )
	{
		XmlElement elem = doc.CreateElement( nodeName );
		elem.InnerText = nodeValue;
		parentElem.AppendChild( elem );
	}


	//--------------------------------------------------------------------------------------------------
	private string PostData( string url, string postData )
	{
		//System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolTypeExtensions.Tls12;
		ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;

		HttpWebRequest request = null;
		Uri uri = new Uri( url );

		request = ( HttpWebRequest )WebRequest.Create( uri );
		request.Method = "POST";
		request.ContentType = "application/x-www-form-urlencoded";

		UTF8Encoding encoding = new UTF8Encoding( );
		byte[] postBytes = encoding.GetBytes( "instream=" + postData );
		request.ContentLength = postBytes.Length;

		Stream requestStream = request.GetRequestStream( );
		requestStream.Write( postBytes, 0, postBytes.Length );
		requestStream.Close( );

		string result = string.Empty;
		HttpWebResponse response = ( HttpWebResponse )request.GetResponse( );

		if ( response.StatusCode != System.Net.HttpStatusCode.OK )
		{
			throw new Exception( "HTTP status code \"" + response.StatusCode.ToString( ) + "\" returned from server" );
		}

		StreamReader reader = new StreamReader( response.GetResponseStream( ), encoding );
		result = reader.ReadToEnd( );
		reader.Close( );
		response.Close( );

		return result;
	}


	//--------------------------------------------------------------------------------------------------
	private Boolean GetPNPOptions( string iOrgId, out string sPartnerCode, out string sUserName, out string sPassword, out string sProductId, out string sOfficeCd )
	{
		int iRecordsRead = 0;

		sPartnerCode = "";
		sUserName = "";
		sPassword = "";
		sProductId = "";
		sOfficeCd = "";

		string sSql = "SELECT partnercode, username, password, productid, officecd ";
		sSql += "FROM egov_pnp_options WHERE orgid = " + iOrgId;

		//Response.Write( "[" + sSql + "]" );

		SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
		sqlConn.Open( );

		SqlCommand myCommand = new SqlCommand( sSql, sqlConn );
		SqlDataReader myReader;
		myReader = myCommand.ExecuteReader( );

		while ( myReader.Read( ) )
		{
			iRecordsRead++;
			sPartnerCode = myReader["partnercode"].ToString( );
			sUserName = myReader["username"].ToString( );
			sPassword = myReader["password"].ToString( );
			sProductId = myReader["productid"].ToString( );
			sOfficeCd = myReader["officecd"].ToString( );
		}

		myReader.Close( );
		sqlConn.Close( );

		if ( iRecordsRead > 0 )
			return true;
		else
			return false;

	}


	//--------------------------------------------------------------------------------------------------
	private Boolean GetPNPURL( string iOrgId, out string sURL )
	{
		int iRecordsRead = 0;

		sURL = "";

		string sSql = "SELECT ISNULL(url,'') AS url FROM egov_pnp_options WHERE orgid = " + iOrgId;

		SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
		sqlConn.Open( );

		SqlCommand myCommand = new SqlCommand( sSql, sqlConn );
		SqlDataReader myReader;
		myReader = myCommand.ExecuteReader( );

		while ( myReader.Read( ) )
		{
			iRecordsRead++;
			sURL = myReader["url"].ToString( );
		}

		myReader.Close( );
		sqlConn.Close( );

		if ( iRecordsRead > 0 )
			return true;
		else
			return false;
	}
    

}
