<!--
var myreq = null;
var myCallBack = '';
var myGetXml = 0;

function createReq() 
{
	try	{
			req = new XMLHttpRequest(); /* FireFox */
		} 
		catch(err1) 
		{
			try
			{
				req = new ActiveXObject("Msxml2.XMLHTTP"); /* some IE */
			}
			catch (err2)
			{
				try
				{
					req = new ActiveXObject("Microsoft.XMLHTTP");  /* some other IE */
				}
				catch (err3)
				{
					alert("Your browser does not support AJAX! Please upgrade to the latest version to access the features of this site.");
					req = false;
				}
			}
		}
		return req;
}

function requestGET( url, query, req )
{
	var myRand = parseInt(Math.random() * 99999999 );
	req.open( "GET", url + '?' + query + '&rand=' + myRand, true );
	req.send(null);
}

function requestPOST( url, query, req )
{
	req.open( "POST", url, true );
	req.setRequestHeader( 'Content-Type', 'application/x-www-form-urlencoded' );
	req.send( query );
}

function doCallback( callback, item )
{
	eval( callback + '(item)' );
}

function doAjax( url, query, callback, reqtype, getxml )
{
	myCallBack = callback;
	myGetXml = getxml;

	// Create the XMLHTTPRequest object instance
	myreq = createReq();

	// set the function to check the state change
	myreq.onreadystatechange = onReadyStateChange; 
	
	// Make the call to the page
	if (reqtype == 'post')
	{
		requestPOST( url, query, myreq );
	}
	else 
	{
		requestGET( url, query, myreq );
	}
}

function onReadyStateChange()
{
	// if completed
	if (myreq.readyState == 4)
	{
		// if successful
		if (myreq.status == 200)
		{
			// capture the returned response
			var item = myreq.responseText;
			if (myGetXml == 1)
			{
				item = myreq.response.XML;
			}
			// Call the function that handles the response
			doCallback(myCallBack, item);
		}
	}
}
//-->