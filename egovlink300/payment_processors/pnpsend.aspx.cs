using System;
using System.Collections;
using System.Configuration;
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


/// <summary>
///--------------------------------------------------------------------------------------------------
///
///--------------------------------------------------------------------------------------------------
/// FILENAME: pnpsend.aspx.cs
/// AUTHOR: Steve Loar
/// CREATED: 07/14/2010
/// COPYRIGHT: Copyright 2010 eclink, inc.
///			 All Rights Reserved.
///
/// DESCTIPTION: This performs the payment transaction with Point and Pay. It is called via HTTP Request from ASP pages. 
///				Needs it's parent page of pnpsend.aspx to work
///
/// MODIFICATION HISTORY
/// 1.0   07/14/2010	Steve Loar - INITIAL VERSION
///
///--------------------------------------------------------------------------------------------------
///
///--------------------------------------------------------------------------------------------------
/// </summary>

public partial class ppv4test_pnpsend : System.Web.UI.Page
{
	//--------------------------------------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
		string sPNPURL;
		string Results = "";
		XmlDocument doc = new XmlDocument( );
		string sPartnerCode;
		string sUserName;
		string sPassWord;
		string sProductId;
		string sOfficeCd;
		string sSVA = generateRequestID( 16 );
		double dChargeAmount;
		string iPaymentControlNumber;
		
		Response.Clear( );

		string sOrgId = common.getOrgId( );

		// Get the passed payment control number
		if ( Request["paymentcontrolnumber"] == "" )
		{
			// log the bad control number
			iPaymentControlNumber = common.CreatePaymentControlRow( sOrgId, "PNP Payment", "Missing Payment Control No" );

			// write that the amount was bad
			Response.Write( "results=0&status=failed&errors=Missing Payment Control number." );
			Response.End( );
		}
		else
			iPaymentControlNumber = Request["paymentcontrolnumber"];

		common.AddToPaymentLog( sOrgId, "PNP Payment", iPaymentControlNumber, "PNP Send Process Started" );

		// bring in the charge Amount passed
		if ( double.TryParse( Request["chargeamount"], out dChargeAmount ) == false )
		{
			// log the bad amount
			common.AddToPaymentLog( sOrgId, "PNP Payment", iPaymentControlNumber, "Amount passed was not valid: " + common.dbSafe( Request["chargeamount"] ) );

			// write that the amount was bad
			Response.Write( "results=0&status=failed&errors=The Charge Amount is not a valid number." );
			Response.End( );
		}

		// Get the PNP options for the account
		if ( GetPNPOptions( sOrgId, out sPartnerCode, out sUserName, out sPassWord, out sProductId, out sOfficeCd ) )
		{
			common.AddToPaymentLog( sOrgId, "PNP Payment", iPaymentControlNumber, "Have Payment Options" );
			XmlElement elemRoot = doc.CreateElement( "root" );
			XmlElement elemData = doc.CreateElement( "data" );
			XmlElement elemStruct = doc.CreateElement( "struct" );
			doc.AppendChild( elemRoot );
			elemRoot.AppendChild( elemData );
			elemData.AppendChild( elemStruct );

			// build the sending xml here
			InsertDataNode( doc, elemStruct, "PartnerCD", sPartnerCode );
			InsertDataNode( doc, elemStruct, "UserName", sUserName );
			InsertDataNode( doc, elemStruct, "Password", sPassWord );
			InsertDataNode( doc, elemStruct, "Action", "PaymentCredit" );
			InsertDataNode( doc, elemStruct, "ProductID", sProductId );
            //if ( sOfficeCd != "" )
                InsertDataNode( doc, elemStruct, "OfficeCD", sOfficeCd );
			InsertDataNode( doc, elemStruct, "ChannelCD", "28" ); // 28 is the API
			InsertDataNode( doc, elemStruct, "SVA", sSVA );
			common.AddToPaymentLog( sOrgId, "PNP Payment", iPaymentControlNumber, "SVA: " + sSVA );
			InsertDataNode( doc, elemStruct, "ChargeAccountNumber", CleanForXML( Request["chargeaccountnumber"] ) );
			InsertDataNode( doc, elemStruct, "ChargeAmount", dChargeAmount.ToString( "#0.00" ) );
			common.AddToPaymentLog( sOrgId, "PNP Payment", iPaymentControlNumber, "ChargeAmount: " + dChargeAmount.ToString( "#0.00" ) );
			InsertDataNode( doc, elemStruct, "ChargeExpirationMMYY", CleanForXML( Request["chargeexpirationmmyy"] ) );
			if ( Request["chargecvn"] != null && Request["chargecvn"] != "" )
			{
				common.AddToPaymentLog( sOrgId, "PNP Payment", iPaymentControlNumber, "CVN: [" + Request["chargecvn"] + "]" );
				InsertDataNode( doc, elemStruct, "ChargeCVN", CleanForXML( Request["chargecvn"] ) );
			}
			InsertDataNode( doc, elemStruct, "SignerFirstName", CleanForXML( Request["signerfirstname"] ) );
			InsertDataNode( doc, elemStruct, "SignerLastName", CleanForXML( Request["signerlastname"] ) );
			common.AddToPaymentLog( sOrgId, "PNP Payment", iPaymentControlNumber, "Signer Name: " + common.dbSafe(Request["signerfirstname"]) + " " + common.dbSafe(Request["signerlastname"]) );
			InsertDataNode( doc, elemStruct, "SignerAddressLine1", CleanForXML( Request["signeraddressline1"] ) );
			InsertDataNode( doc, elemStruct, "SignerAddressCity", CleanForXML( Request["signeraddresscity"] ) );
			InsertDataNode( doc, elemStruct, "SignerAddressRegionCode", CleanForXML( Request["signeraddressregioncode"] ) );
			InsertDataNode( doc, elemStruct, "SignerAddressPostalCode", CleanForXML( Request["signeraddresspostalcode"] ) );
			InsertDataNode( doc, elemStruct, "SignerAddressCountryCode", "US" );
			InsertDataNode( doc, elemStruct, "Notes", CleanForXML( Request["notes"] ) );
			common.AddToPaymentLog( sOrgId, "PNP Payment", iPaymentControlNumber, "Notes: " + Request["notes"] );

			// turn the xml into a string for sending
			string requestStr = doc.OuterXml;

//			Response.Write( requestStr );
//			Response.End( );

			// fetch the PNP url here and send the request
			if ( GetPNPURL( sOrgId, out sPNPURL ) == true )
			{
				common.AddToPaymentLog( sOrgId, "PNP Payment", iPaymentControlNumber, "Posting data to PNP" );
				Results = PostData( sPNPURL, requestStr );
				common.AddToPaymentLog( sOrgId, "PNP Payment", iPaymentControlNumber, "Returned from PNP" );
			}
			else
			{
				Response.Write( "results=0&status=failed&errors=Failed to get payment processor URL." );
				Response.End( );
			}

			// capture the PNP results in XML
			XmlDocument resultsDoc = new XmlDocument( );
			resultsDoc.LoadXml( Results );
			XmlElement root = resultsDoc.DocumentElement;

			XmlNodeList oNodes = root.FirstChild.FirstChild.ChildNodes;

			// start our send back string with the results count and our sva value
			string sResults = "result=" + oNodes.Count.ToString( );
			sResults += "&sva=" + sSVA;

			// buld the string from the returned XML nodes. Should be order number, payment, fee, total, created, errors, status.
			foreach ( XmlNode node in oNodes )
			{
				sResults += "&" + node.Name.ToLower( ) + "=" + node.InnerText;
			}

			// if we did not get a status back, then we failed some test and so are declined
			if ( sResults.IndexOf( "status" ) < 1 )
				sResults += "&status=Declined";

			common.AddToPaymentLog( sOrgId, "PNP Payment", iPaymentControlNumber, "Results: " + common.dbSafe( sResults ) );

			// Send back the results to the calling code
			Response.Write( sResults );
			Response.End();
		}
		else
		{
			// write that the option fetch failed
			Response.Write( "results=0&status=failed&errors=Failed to get payment options." );
			Response.End( );
		}
    }


	//--------------------------------------------------------------------------------------------------
	private string CleanForXML( string _sValue )
	{
		string sValue = _sValue.Replace( "&", "&amp;" );
		sValue = sValue.Replace( ">", "&gt;" );
		sValue = sValue.Replace( "<", "&lt;" );
		//sValue = sValue.Replace( "'", "&#39;" );
        sValue = sValue.Replace( "'", "&apos;" );
		sValue = sValue.Replace( "\"", "&quot;" );

		return sValue;
	}


	//--------------------------------------------------------------------------------------------------
	private string generateRequestID( int iGUIDLength )
	{
		const string strValid = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";

		Random r = new Random( );
		char[] charArray = new char[iGUIDLength];
		string charPool = string.Empty;
		int index;

		//Build the output character array
		for ( int i = 0; i < charArray.Length; i++ )
		{
			//Pick a random integer in the character pool
			index = r.Next( 0, strValid.Length );

			//Set it to the output character array
			charArray[i] = strValid[index];
		}

		return new string( charArray );
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
            //int officeCD;
            //int.TryParse( myReader["officecd"].ToString( ), out officeCD );
            //if ( officeCD < 0 )
            //    sOfficeCd = "";
            //else
                sOfficeCd = myReader["officecd"].ToString( );
		}

		myReader.Close( );
		sqlConn.Close( );

		if ( iRecordsRead > 0 )
			return true;
		else
			return false;

	}

}
