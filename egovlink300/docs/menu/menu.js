
/* TOC.JS */
var framesTop = parent;
var L_LoadingMsg_HTMLText = "Loading, click to cancel...";
var LoadDiv = '<DIV ONCLICK="loadFrame(true);" CLASS="clsLoadMsg">';

L_LoadingMsg_HTMLText = LoadDiv + L_LoadingMsg_HTMLText + "</LI>";

var currentPath = "";
var ob, upOb, X, Y, wasDrag, dragNumLevels;

varContextOn = true;

function RefreshPath() {
  document.location.reload();

  /*var eLI = document.getElementById( currentPath );
  MarkActive(eLI);
  var eUL = GetNextUL(eLI);

  bLoading = true;
  eLI.className = "kidShown";
  eUL.className = "clsShown";
  window.eCurrentUL = eUL;
  eUL.innerHTML = L_LoadingMsg_HTMLText;
  document.frames["hiddenframe"].location.replace("menu/loadtree.asp?path=" + currentPath);*/
}


//------- DRAG AND DROP FUNCTIONS -------------------------
function MD(e) {
  ob = event.srcElement.parentElement;
  X = event.clientX;
  Y = event.clientY;
  ob.style.cursor = "move";
}
function MM(e) {
  if (ob != null) {
    ob.children[0].style.filter = "alpha(opacity=30)";
    ob.style.pixelLeft = event.clientX - X;
    ob.style.pixelTop = event.clientY - Y;
    return false;
  }
}
function MU() {
  var spath, sname, dpath, dname;
  
  wasDrag = true;
  if (ob == null) {
    wasDrag = false;
    return;
  }

  if (ob.style.pixelTop == 0 || ob.style.pixelLeft == 0)
    wasDrag = false;
  
  ob.style.pixelLeft = 0;
  ob.style.pixelTop = 0;
  ob.children[0].style.filter = "alpha(opacity=100)";
  ob.style.cursor = "default";
  
  if (wasDrag) {
    if (lastObj != null && lastObj != ob && lastObj.nodeType == "c") {
      spath = ob.id;
      sname = getPathObject(spath);
      dpath = lastObj.id;
      dname = getPathObject(dpath);
      if (ob.nodeType == "c") {
        if (confirm("Do you really want to move the category \"" + sname + "\" to \"" + dname + "\"?")) {
          parent.fraTopic.document.location.href = "../moveobj.asp?t=c&s=" + spath + "&d=" + dpath + "/" + sname;
        }
      }
      else if (ob.nodeType == "a") {
        if (confirm("Do you really want to move the article \"" + sname + "\" to the category \"" + dname + "\"?"))
          parent.fraTopic.document.location.href = "../moveobj.asp?t=a&s=" + spath + "&d=" + dpath + "/" + sname;
      }
      wasDrag = false;
      ob = null;
    }
    else {
      upOb = ob;
    }
  }
  ob = null;
}
function getPathObject( path ) {
  pos = path.indexOf('/');
  while (pos >= 0) {
    path = path.substr(pos+1,path.length-pos+1);
    pos = path.indexOf('/');
  }
  pos = path.indexOf("%2E");
  while (pos >=0) {
    path = path.substr(0,pos) + "." + path.substr(pos+3,path.length-pos+3);
    pos = path.indexOf("%2E");
  }
  pos = path.indexOf("%20");
  while (pos >=0) {
    path = path.substr(0,pos) + " " + path.substr(pos+3,path.length-pos+3);
    pos = path.indexOf("%20");
  }
  pos = path.indexOf("+");
  while (pos >=0) {
    path = path.substr(0,pos) + " " + path.substr(pos+1,path.length-pos+1);
    pos = path.indexOf("+");
  }

  return path;
}
document.onmousedown  = MD;
//document.onmousemove  = MM;
document.onmouseup    = MU;

//-------- END DRAG AND DROP --------------------------------


// -----------------------------------------------------------
	// Client-side BrowserData constructor
	// Populated using data from server-side oBD object to avoid redundancy
	// -----------------------------------------------------------
	function BrowserData()
	{
		this.userAgent = "Mozilla/4.0 (compatible; MSIE 5.5; Windows 98)";
		this.browser = "MSIE";
		this.majorVer = 5;
		this.minorVer = "0";
		this.betaVer = "0";
		this.platform = "98";
		this.doesDHTML = true;
		this.doesActiveX = true;
	}
	var oBD = new BrowserData();
	
	
function NoOp() {
    return;
}
  
function caps(){
    var UA = navigator.userAgent;
    if(UA.indexOf("MSIE") != -1){
        this.ie = true;
        var v = UA.charAt(UA.indexOf("MSIE") + 5);
        if(v == 2 ) this.ie2 = true;
        else if(v == 3 ) this.ie3 = true;
        else if(v == 4 ) this.ie4 = true;
        else if(v == 5 ) this.ie5 = true;
        if(this.ie4 || this.ie5) this.UL = true;
    }else if(UA.indexOf("Mozilla") != -1 && UA.indexOf("compatible") == -1){
        this.nav = true;
        var v = UA.charAt(UA.indexOf("Mozilla") + 8);
        if(v == 2 ) this.nav2 = true;
        else if(v == 3 ) this.nav3 = true;
        else if(v == 4 ) this.nav4 = true;
    }
    if(UA.indexOf("Windows 95") != -1 || UA.indexOf("Win95") != -1 || UA.indexOf("Win98") != -1 || UA.indexOf("Windows 98") != -1 || UA.indexOf("Windows NT") != -1) this.win32 = true;
    else if(UA.indexOf("Windows 3.1") != -1 || UA.indexOf("Win16") != -1) this.win16 = true;
    else if(UA.indexOf("Mac") != -1) this.anymac = true;
    else if(UA.indexOf("SunOS") != -1 || UA.indexOf("HP-UX") != -1 || UA.indexOf("X11") != -1) this.unix = true;
    else if(UA.indexOf("Windows CE") != -1) this.wince = true;
}

var bc = new caps();

////////////////////////////////////////////
// Not sure why this is here, it puts a scrollbar up when none is needed
// if("object" == typeof(parent.document.all.fraPaneToc)) parent.document.all.fraPaneToc.scrolling = "yes";
////////////////////////////////////////////

var eSynchedNode = null;
var eCurrentUL = null;
var eCurrentLI = null;
var bLoading = false;

function loadFrame( bStopLoad )
{
    if( "object" == typeof( eCurrentUL ) && eCurrentUL && !bStopLoad )
    {
      //window.scrollTo(0,eCurrentUL.offsetTop-(document.body.clientHeight/2));
      
      eCurrentUL.innerHTML = hiddenframe.chunk.innerHTML;
      eCurrentUL = null;
      bLoading = false;
    }
    else if( "object" == typeof( eCurrentUL ) && eCurrentUL )
    {
      eCurrentUL.parentElement.children[1].className = "";
      if (eSrc.src.indexOf("locked") <=0)
        eCurrentUL.parentElement.children[0].src = "images/folder_closed.gif";
      else
       eCurrentUL.parentElement.children[0].src = "images/locked_folder_closed.gif";
      //eCurrentUL.parentElement.children[0].src = "images/folder_closed.gif";
      eCurrentUL.parentElement.className = "kid";
      eCurrentUL.className = "clsHidden";
      eCurrentUL.innerHTML="";
      eCurrentUL = null;
      bLoading = false;
    }
    else
    {
      bLoading = false;
    }
    return;
}

function GetNextUL(eSrc)
{
    var eRef = eSrc;
    for(var i = 0; i < eRef.children.length; i++) if("UL" == eRef.children[i].tagName) return eRef.children[i];
    return false;
}

function MarkSync(eSrc)
{
    if("object" == typeof(aNodeTree)) aNodeTree = null;
    if("LI" == eSrc.tagName.toUpperCase() && eSrc.children[1] && eSynchedNode != eSrc )
    {
        UnmarkSync();
        eSrc.children[1].style.fontWeight = "bold";
        eSynchedNode = eSrc;
	}
}

function UnmarkSync()
{
    if("object" == typeof(eSynchedNode) && eSynchedNode )
    {
        eSynchedNode.children[1].style.fontWeight = "normal";
        eSynchedNode = null;
    }
}

function MarkActive(eLI)
{
    if( "object" == typeof( eLI ) && eLI && "LI" == eLI.tagName.toUpperCase() && eLI.children[1] && eLI != eCurrentLI )
    {
        MarkInActive();
        window.eCurrentLI = eLI;
        window.eCurrentLI.children[1].className = "clsCurrentLI";
    }
}

function MarkInActive()
{
    if( "object" == typeof( eCurrentLI ) && eCurrentLI )
    {
        window.eCurrentLI.children[1].className = "";
        window.eCurrentLI = null;
    }
}

function Navigate_URL( eSrc )
{

    var eLink = eSrc.parentElement.children[1];
    urlIdx = eLink.href.indexOf( "URL=" );
	if("object" == typeof(framesTop.fraTopic) && eLink && "A" == eLink.tagName && urlIdx != -1 && "fraTopic" != eLink.target)
    {

        if(eLink.target=="fraTopic"||eLink.target=="_top"){
			
            parent.fraTopic.location.href = eSrc.parentElement.children[1].href;
			//framesTop.fraTopic.location.href = eSrc.parentElement.children[1].href.substring( urlIdx + 4 );
        }else{
			
            window.open(eSrc.parentElement.children[1].href,eLink.target);
        }
		
        MarkSync(eSrc.parentElement);
    }
    else if("object" == typeof(framesTop.fraTopic) && eLink && "A" == eLink.tagName  && eLink.href.indexOf( "tocPath=" ) == -1 && eLink.href.indexOf( "javascript:" ) == -1 )
    {
				
        if(eLink.target=="fraTopic")
        {
			
            parent.fraTopic.location.href = eSrc.parentElement.children[1].href;
			//framesTop.fraTopic.location.href = eSrc.parentElement.children[1].href;
        }
        else if( eLink.target=="_top" )
        {
			
            top.location = eLink.href;
            return;
        }
        else
        {
			
			//window.open(eSrc.parentElement.children[1].href,eLink.target);
        }
		
        MarkSync(eSrc.parentElement);
    }
    else if( eSynchedNode != eSrc.parentElement && ( urlIdx != -1 || ( eLink.href.indexOf( "javascript:" ) == -1 && eLink.href.indexOf( "tocPath=" ) == -1 ) ) )
    {


		if (eSrc.parentElement.nodeType=="a")
		{
			window.open(eSrc.parentElement.children[1].href,"_NEW");
		}
				
		MarkSync( eSrc.parentElement );
    }
}

function Image_Click( eSrc , bLeaveOpen )
{
    var eLink = eSrc.parentElement.children[1];
    if("noHand" != eSrc.className)
    {
        eLI = eSrc.parentElement;
        MarkActive(eLI);
        var eUL = GetNextUL(eLI);
        if(eUL && "kidShown" == eLI.className)
        {
            // hide on-page kids
            if( !bLeaveOpen )
            {
                eLI.className = "kid";
                eUL.className = "clsHidden";
                if (eSrc.src.indexOf("images/help.gif") <= 0) {
                  if (eSrc.src.indexOf("locked") <=0)
                    eSrc.src = "images/folder_closed.gif";
                  else
                    eSrc.src = "images/locked_folder_closed.gif";
                  //eSrc.src = "images/folder_closed.gif";
                }
            }
        }
        else if(eUL && eUL.all.length && "kid" == eLI.className)
        {
            // show on-page kids
            eLI.className = "kidShown";
            eUL.className = "clsShown";
            if (eSrc.src.indexOf("images/help.gif") <= 0) {
              if (eSrc.src.indexOf("locked") <=0)
                eSrc.src = "images/folder_open.gif";
              else
                eSrc.src = "images/locked_folder_open.gif";
            }
        }
        else if("kid" == eLI.className)
        {
            // load off-page kids
            if( !bLoading )
            {
                bLoading = true;         
                eLI.className = "kidShown";
                eUL.className = "clsShown";
                window.eCurrentUL = eUL;
                if (eSrc.src.indexOf("images/help.gif") <= 0) {
                  if (eSrc.src.indexOf("locked") <=0)
                    eSrc.src = "images/folder_open.gif";
                  else
                    eSrc.src = "images/locked_folder_open.gif";
                }
                //eUL.innerHTML = L_LoadingMsg_HTMLText;
				        var strLoc = "loadtree.asp" + eLink.href.substring( eLink.href.indexOf( "?" ));
                document.frames["hiddenframe"].location.replace(strLoc);

                //this is so we can only refresh one location in the directory tree
                currentPath = eLink.parentNode.id;
            }
        }
    }
}

function syncTo(nodeName) {
	var eSrc = eval("document.all." + nodeName);
    //event.returnValue = false;

    if("A" == eSrc.tagName.toUpperCase() && "LI" == eSrc.parentElement.tagName)
    {
        var eImg = eSrc.parentElement.children[0];
        if(eImg) eImg.click();
    }
    else if("SPAN" == eSrc.tagName && "LI" == eSrc.parentElement.tagName)
    {
        var eImg = eSrc.parentElement.children[0];
        if(eImg) eImg.click();
    }
    else if("IMG" == eSrc.tagName)
    {
        Image_Click( eSrc , false );
        Navigate_URL( eSrc );
    }
    return event.returnValue;
}

function Toc_click()
{
    if (wasDrag)
      return false;
    
    var eSrc = window.event.srcElement;
    event.returnValue = false;

    if("A" == eSrc.tagName.toUpperCase() && "LI" == eSrc.parentElement.tagName)
    {
        var eImg = eSrc.parentElement.children[0];
        if(eImg) eImg.click();
    }
    else if("SPAN" == eSrc.tagName && "LI" == eSrc.parentElement.tagName)
    {
        var eImg = eSrc.parentElement.children[0];
        if(eImg) eImg.click();
    }
    else if("IMG" == eSrc.tagName)
    {
        Image_Click( eSrc , false );
		Navigate_URL( eSrc );
    }
    return event.returnValue;
}

function window_load()
{
    //if( self == top ) location.replace( "../default.asp" );
    var objStyle = null;
    if( "MSIE" == oBD.browser && 3 < oBD.majorVer && "Mac" != oBD.platform && "object" == typeof ( ulRoot ) && "object" == typeof( objStyle = document.styleSheets[0] ) && "object" == typeof( objStyle.addRule ) )
    {
        window.eSynchedNode = document.all["eSynchedNode"];
        objStyle.addRule( "UL.clsHidden" , "display:none" , 0 );
        objStyle.addRule( "UL.hdn" , "display:none" , 0 );
        ulRoot.onclick=Toc_click;
        /*if( window.eSynchedNode )
        {
            MarkActive(window.eSynchedNode);
            window.eSynchedNode.all.tags( "B" )[0].outerHTML = eSynchedNode.all.tags("B")[0].innerHTML;
            window.scrollTo(0,window.eSynchedNode.offsetTop-(document.body.clientHeight/2));
        }
        else
        {
            MarkActive(document.all.tags( "LI" )[0]);
        }*/
    }
}

window.onload = window_load;

//-------- CONTEXT MENU STUFF--------------------------------------------
  var lastObj = null;
  var actOnObj = null;
  var display_url = false;
  var ie5 = document.all && document.getElementById;
  var ns6 = document.getElementById && !document.all;
  var menuobj = null;
  var last_menuobj = null;

  function mover(){
    lastObj = window.event.srcElement;
    if (lastObj.tagName.toUpperCase() != "LI")
      lastObj = lastObj.parentElement;
    window.event.cancelBubble;

    if (lastObj != null) {
      if (wasDrag) {
        //handle if drag upwards
        if (lastObj != null && lastObj != upOb && lastObj.nodeType == "c") {
          spath = upOb.id;
          sname = getPathObject(spath);
          dpath = lastObj.id;
          dname = getPathObject(dpath);
          if (upOb.nodeType == "c") {
            if (confirm("Do you really want to move the category \"" + sname + "\" to \"" + dname + "\"?")) {
              parent.fraTopic.document.location.href = "../moveobj.asp?t=c&s=" + spath + "&d=" + dpath + "/" + sname;
            }
          }
          else if (upOb.nodeType == "a") {
            if (confirm("Do you really want to move the article \"" + sname + "\" to the category \"" + dname + "\"?"))
              parent.fraTopic.document.location.href = "../moveobj.asp?t=a&s=" + spath + "&d=" + dpath + "/" + sname;
          }
        }
        wasDrag = false;
        upOb = null;
      }
    }
  }
  function mout(){
    lastObj = null;
  }
  document.onmouseover = mover;
  document.onmouseout = mout;

  function initContextMenu(){
    if (ie5||ns6) {
      varContextOn = true;
      document.oncontextmenu = showmenuie5;
      document.onclick = hidemenuie5;
    }
  }

  function showmenuie5(e){
    if (lastObj != null)
      actOnObj = lastObj;
    else
      return false;

    hidemenuie5();
    if (actOnObj.nodeType == "c") 
      menuobj = document.getElementById("mnu_Category");
    else
      menuobj = document.getElementById("mnu_Article");
      
    //Find out how close the mouse is to the corner of the window
    var rightedge=ie5? document.body.clientWidth-event.clientX : window.innerWidth-e.clientX
    var bottomedge=ie5? document.body.clientHeight-event.clientY : window.innerHeight-e.clientY

    //if the horizontal distance isn't enough to accomodate the width of the context menu
    if (rightedge<menuobj.offsetWidth)
      menuobj.style.left=ie5? document.body.scrollLeft+event.clientX-menuobj.offsetWidth : window.pageXOffset+e.clientX-menuobj.offsetWidth
    else
      menuobj.style.left=ie5? document.body.scrollLeft+event.clientX : window.pageXOffset+e.clientX

    //same concept with the vertical position
    if (bottomedge<menuobj.offsetHeight)
      menuobj.style.top=ie5? document.body.scrollTop+event.clientY-menuobj.offsetHeight : window.pageYOffset+e.clientY-menuobj.offsetHeight
    else
      menuobj.style.top=ie5? document.body.scrollTop+event.clientY : window.pageYOffset+e.clientY

    menuobj.style.visibility = "visible";
    menuobj.style.display = "";
    last_menuobj = menuobj;
    return false;
  }

  function hidemenuie5(e){
    if (last_menuobj != null) {
      last_menuobj.style.visibility = "hidden";
      last_menuobj.style.display = "none";
    }
    last_menuobj = null;
  }

  function highlightie5(e){
    var firingobj=ie5? event.srcElement : e.target;
    if (firingobj.className == "menuitems" || firingobj.className == "menuimage" || ns6&&firingobj.parentNode.className == "menuitems"){
      if (ns6&&firingobj.parentNode.className=="menuitems")
        firingobj=firingobj.parentNode; //up one node
      if (firingobj.className == "menuimage")
        firingobj=firingobj.parentNode;
      
      firingobj.style.backgroundColor="#336699";
      firingobj.style.color="white";
      
      if (display_url)
        window.status=event.srcElement.url;
    }
  }

  function lowlightie5(e){
    var firingobj=ie5? event.srcElement : e.target;
    if (firingobj.className=="menuitems" || firingobj.className == "menuimage" || ns6&&firingobj.parentNode.className=="menuitems"){
      if (ns6&&firingobj.parentNode.className=="menuitems")
        firingobj=firingobj.parentNode; //up one node
      if (firingobj.className == "menuimage")
        firingobj=firingobj.parentNode;

      firingobj.style.backgroundColor="";
      firingobj.style.color="black";
      window.status='';
    }
  }

  function jumptoie5(e){
    var firingobj=ie5? event.srcElement : e.target;

    if (firingobj.className == "menuitems" || firingobj.className == "menuimage" || ns6&&firingobj.parentNode.className == "menuitems"){
      if (ns6&&firingobj.parentNode.className == "menuitems")
        firingobj = firingobj.parentNode;
      if (firingobj.className == "menuimage")
        firingobj=firingobj.parentNode;

      parent.fraTopic.window.location = firingobj.getAttribute("url") + "?path=" + actOnObj.id;
    }
  }
  
  
  function ToggleContextMenu(){
    if (ie5||ns6) {
          if (varContextOn == true){
            varContextOn = false;
            hidemenuie5();
            document.oncontextmenu = passthru;}
          else{
			initContextMenu();
          }
    }
  }

  function passthru() {
    
    bStopLoad = false;
    eSrc = window.event.srcElement;
    eLI =  eSrc.parentElement;
    MarkActive(eLI);

   /* if (eSrc.src.indexOf("locked") <=0)
      eSrc.src = "images/folder_open.gif";
    else
      eSrc.src = "images/locked_folder_open.gif";*/

    var eUL = GetNextUL(eLI);
    eCurrentUL = eUL;

    //eLI.className = "kidShown";
    eUL.className = "clsShown";
    
    return true;
  }
  
//------- END CONTEXT MENU STUFF-----------------------------
//-->