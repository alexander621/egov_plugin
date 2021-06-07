// Title: tigra menu
// Description: See the demo at url
// URL: http://www.javascript-menu.com/
// Version: 2.0 (commented source)
// Date: 04-05-2003 (mm-dd-yyyy)
// Tech. Support: http://www.softcomplex.com/forum/forumdisplay.php?fid=40
// Notes: This script is free. Visit official site for further details.

// --------------------------------------------------------------------------------
// global collection containing all menus on current page
var A_MENUS = [];

// --------------------------------------------------------------------------------
// menu class
function menu (a_items, a_tpl) {

	// browser check
	if (!document.body || !document.body.style)
		return;

	// store items structure
	this.a_config = a_items;

	// store template structure
	this.a_tpl = a_tpl;

	// get menu id
	this.n_id = A_MENUS.length;

	// declare collections
	this.a_index = [];
	this.a_children = [];

	// assigh methods and event handlers
	this.expand      = menu_expand;
	this.collapse    = menu_collapse;

	this.onclick     = menu_onclick;
	this.onmouseout  = menu_onmouseout;
	this.onmouseover = menu_onmouseover;
	this.onmousedown = menu_onmousedown;

	// default level scope description structure 
	this.a_tpl_def = {
		'block_top'  : 16,
		'block_left' : 16,
		'top'        : 20,
		'left'       : 4,
		'width'      : 120,
		'height'     : 22,
		'hide_delay' : 0,
		'expd_delay' : 0,
		'css'        : {
			'inner' : '',
			'outer' : ''
		}
	};
	
	// assign methods and properties required to imulate parent item
	this.getprop = function (s_key) {
		return this.a_tpl_def[s_key];
	};

	this.o_root = this;
	this.n_depth = -1;
	this.n_x = 0;
	this.n_y = 0;

	// 	init items recursively
	for (n_order = 0; n_order < a_items.length; n_order++)
		new menu_item(this, n_order);

	// register self in global collection
	A_MENUS[this.n_id] = this;

	// make root level visible
	for (var n_order = 0; n_order < this.a_children.length; n_order++)
		this.a_children[n_order].e_oelement.style.visibility = 'visible';
}

// --------------------------------------------------------------------------------
function menu_collapse (n_id) {
	// cancel item open delay
	clearTimeout(this.o_showtimer);

	// by default collapse to root level
	var n_tolevel = (n_id ? this.a_index[n_id].n_depth : 0);
	
	// hide all items over the level specified
	for (n_id = 0; n_id < this.a_index.length; n_id++) {
		var o_curritem = this.a_index[n_id];
		if (o_curritem.n_depth > n_tolevel && o_curritem.b_visible) {
			o_curritem.e_oelement.style.visibility = 'hidden';
			o_curritem.b_visible = false;
		}
	}

	// reset current item if mouse has gone out of items
	if (!n_id)
		this.o_current = null;
}

// --------------------------------------------------------------------------------
function menu_expand (n_id) {

	// expand only when mouse is over some menu item
	if (this.o_hidetimer)
		return;

	// lookup current item
	var o_item = this.a_index[n_id];

	// close previously opened items
	if (this.o_current && this.o_current.n_depth >= o_item.n_depth)
		this.collapse(o_item.n_id);
	this.o_current = o_item;

	// exit if there are no children to open
	if (!o_item.a_children)
		return;

	// show direct child items
	for (var n_order = 0; n_order < o_item.a_children.length; n_order++) {
		var o_curritem = o_item.a_children[n_order];
		o_curritem.e_oelement.style.visibility = 'visible';
		o_curritem.b_visible = true;
	}
}

// --------------------------------------------------------------------------------
//
// --------------------------------------------------------------------------------
function menu_onclick (n_id) {
	// don't go anywhere if item has no link defined
	return Boolean(this.a_index[n_id].a_config[1]);
}

// --------------------------------------------------------------------------------
function menu_onmouseout (n_id) {

	// lookup new item's object	
	var o_item = this.a_index[n_id];

	// apply rollout
	o_item.e_oelement.className = o_item.getstyle(0, 0);
	o_item.e_ielement.className = o_item.getstyle(1, 0);
	
	// update status line	
	//o_item.upstatus(7);
	

	// run mouseover timer
	this.o_hidetimer = setTimeout('A_MENUS['+ this.n_id +'].collapse();',
		o_item.getprop('hide_delay'));

	o_item.unhidethings();
}

// --------------------------------------------------------------------------------
function menu_onmouseover (n_id) {

	// cancel mouseoute menu close and item open delay
	clearTimeout(this.o_hidetimer);
	this.o_hidetimer = null;
	clearTimeout(this.o_showtimer);

	// lookup new item's object	
	var o_item = this.a_index[n_id];

	// update status line	
	//o_item.upstatus();

	// apply rollover
	o_item.e_oelement.className = o_item.getstyle(0, 1);
	o_item.e_ielement.className = o_item.getstyle(1, 1);
	
	// if onclick open is set then no more actions required
	if (o_item.getprop('expd_delay') < 0)
		return;

	// run expand timer
	this.o_showtimer = setTimeout('A_MENUS['+ this.n_id +'].expand(' + n_id + ');',
		o_item.getprop('expd_delay'));

	o_item.hidethings();
}


// --------------------------------------------------------------------------------
// called when mouse button is pressed on menu item
// --------------------------------------------------------------------------------
function menu_onmousedown (n_id) {
	
	// lookup new item's object	
	var o_item = this.a_index[n_id];

	// apply mouse down style
	o_item.e_oelement.className = o_item.getstyle(0, 2);
	o_item.e_ielement.className = o_item.getstyle(1, 2);

	this.expand(n_id);
//	this.items[id].switch_style('onmousedown');
}


// --------------------------------------------------------------------------------
// menu item Class
function menu_item (o_parent, n_order) {

	// store parameters passed to the constructor
	this.n_depth  = o_parent.n_depth + 1;
	this.a_config = o_parent.a_config[n_order + (this.n_depth ? 3 : 0)];

	// return if required parameters are missing
	if (!this.a_config) return;

	// store info from parent item
	this.o_root    = o_parent.o_root;
	this.o_parent  = o_parent;
	this.n_order   = n_order;

	// register in global and parent's collections
	this.n_id = this.o_root.a_index.length;
	this.o_root.a_index[this.n_id] = this;
	o_parent.a_children[n_order] = this;

	// calculate item's coordinates
	var o_root = this.o_root,
		a_tpl  = this.o_root.a_tpl;

	// assign methods
	this.getprop  = mitem_getprop;
	this.getstyle = mitem_getstyle;
	this.upstatus = mitem_upstatus;
	this.hidethings = mitem_hidethings;
	this.unhidethings = mitem_unhidethings;

	this.n_x = n_order
		? o_parent.a_children[n_order - 1].n_x + this.getprop('left')
		: o_parent.n_x + this.getprop('block_left');

	this.n_y = n_order
		? o_parent.a_children[n_order - 1].n_y + this.getprop('top')
		: o_parent.n_y + this.getprop('block_top');

	// generate item's HMTL
	//+ (this.a_config[2] && this.a_config[2]['tw'] ? ' 
	//+ this.a_config[2]['tw'] + '"' : '')
	//	Modified by SJL to use target=_top 10/16/2006
	document.write (
		'<a id="e' + o_root.n_id + '_'
			+ this.n_id +'o" class="' + this.getstyle(0, 0) + '" href="' + this.a_config[1] + '"target="_top"'
			+ (this.a_config[2] && this.a_config[2]['tt'] ?
			+ ' title="'
			+ this.a_config[2]['tt'] + '"' : '') + ' style="position: absolute; top: '
			+ this.n_y + 'px; left: ' + this.n_x + 'px; width: '
			+ this.getprop('width') + 'px; height: '
			+ this.getprop('height') + 'px; visibility: hidden;'
			+' z-index: ' + this.n_depth + ';" '
			+ 'onclick="return A_MENUS[' + o_root.n_id + '].onclick('
			+ this.n_id + ');" onmouseout="A_MENUS[' + o_root.n_id + '].onmouseout('
			+ this.n_id + ');" onmouseover="A_MENUS[' + o_root.n_id + '].onmouseover('
			+ this.n_id + ');" onmousedown="A_MENUS[' + o_root.n_id + '].onmousedown('
			+ this.n_id + ');"><div  id="e' + o_root.n_id + '_'
			+ this.n_id +'i" class="' + this.getstyle(1, 0) + '">'
			+ this.a_config[0] + "</div></a>\n"
		);
	this.e_ielement = document.getElementById('e' + o_root.n_id + '_' + this.n_id + 'i');
	this.e_oelement = document.getElementById('e' + o_root.n_id + '_' + this.n_id + 'o');

	this.b_visible = !this.n_depth;

	// no more initialization if leaf
	if (this.a_config.length < 4)
		return;

	// node specific methods and properties
	this.a_children = [];

	// init downline recursively
	for (var n_order = 0; n_order < this.a_config.length - 3; n_order++)
		new menu_item(this, n_order);

}

// --------------------------------------------------------------------------------
// reads property from template file, inherits from parent level if not found
// ------------------------------------------------------------------------------------------
function mitem_getprop (s_key) {

	// check if value is defined for current level
	var s_value = null,
		a_level = this.o_root.a_tpl[this.n_depth];

	// return value if explicitly defined
	if (a_level)
		s_value = a_level[s_key];

	// request recursively from parent levels if not defined
	return (s_value == null ? this.o_parent.getprop(s_key) : s_value);
}
// --------------------------------------------------------------------------------
// reads property from template file, inherits from parent level if not found
// ------------------------------------------------------------------------------------------
function mitem_getstyle (n_pos, n_state) {

	var a_css = this.getprop('css');
	var a_oclass = a_css[n_pos ? 'inner' : 'outer'];

	// same class for all states	
	if (typeof(a_oclass) == 'string')
		return a_oclass;

	// inherit class from previous state if not explicitly defined
	for (var n_currst = n_state; n_currst >= 0; n_currst--)
		if (a_oclass[n_currst])
			return a_oclass[n_currst];
}

// ------------------------------------------------------------------------------------------
// updates status bar message of the browser
// ------------------------------------------------------------------------------------------
function mitem_upstatus (b_clear) {
	window.setTimeout("window.status=unescape('" + (b_clear
		? ''
		: (this.a_config[2] && this.a_config[2]['sb']
			? escape(this.a_config[2]['sb'])
			: escape(this.a_config[0]) + (this.a_config[1]
				? ' ('+ escape(this.a_config[1]) + ')'
				: ''))) + "')", 10);
}


function mitem_hidethings()
{
		// This form is found on security/edit_user_security.asp and security/copy_user_security.asp
		var formnames = document.getElementsByName("UserForm");
		if (formnames.length == 1)
		{
			var bexists = eval(document.UserForm["iUserID"]);
			if(bexists)
			{
				document.UserForm.iUserID.style.visibility="hidden";
			}
			bexists = eval(document.UserForm["iFromUserID"]);
			if(bexists)
			{
				document.UserForm.iFromUserID.style.visibility="hidden";
			}
			bexists = eval(document.UserForm["iToUserID"]);
			if(bexists)
			{
				document.UserForm.iToUserID.style.visibility="hidden";
			}
		}
		// This form is found on admin/manage_features.asp
		var formnames = document.getElementsByName("pickForm");
		if (formnames.length == 1)
		{
			var bexists = eval(document.pickForm["orgid"]);
			if(bexists)
			{
				document.pickForm.orgid.style.visibility="hidden";
			}
		}
		// This form is found on poolpass/poolpass_form.asp
		var formnames = document.getElementsByName("BuyerForm");
		if (formnames.length == 1)
		{
			var bexists = eval(document.BuyerForm["userid"]);
			if(bexists)
			{
				document.BuyerForm.userid.style.visibility="hidden";
			}
		}
		// This form is found on poolpass/poolpass_rates.asp
		var formnames = document.getElementsByName("rateform0");
		if (formnames.length == 1)
		{
			var bexists = eval(document.rateform0["iPeriodId"]);
			if(bexists)
			{
				document.rateform0.iPeriodId.style.visibility="hidden";
			}
		}
		// This form is found on poolpass/poolpass_rates.asp
		var formnames = document.getElementsByName("rateform1");
		if (formnames.length == 1)
		{
			var bexists = eval(document.rateform1["iPeriodId"]);
			if(bexists)
			{
				document.rateform1.iPeriodId.style.visibility="hidden";
			}
		}
		// This form is found on poolpass/poolpass_type_report.asp, classes/class_statisticsreport.asp
		var formnames = document.getElementsByName("YearForm");
		if (formnames.length == 1)
		{
			var bexists = eval(document.YearForm["iyear"]);
			if(bexists)
			{
				document.YearForm.iyear.style.visibility="hidden";
			}
		}
		// This form is found on gifts/gift_form.asp
		var formnames = document.getElementsByName("frmpayment");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmpayment["gift"]);
			if(bexists)
			{
				document.frmpayment.gift.style.visibility="hidden";
			}
			bexists = eval(document.frmpayment["userid"]);
			if(bexists)
			{
				document.frmpayment.userid.style.visibility="hidden";
			}
		}
		// This form is found on recreation/facility_calendar.asp
		var formnames = document.getElementsByName("frmcal");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmcal["selfacility"]);
			if(bexists)
			{
				document.frmcal.selfacility.style.visibility="hidden";
			}
			bexists = eval(document.frmcal["selmonth"]);
			if(bexists)
			{
				document.frmcal.selmonth.style.visibility="hidden";
			}
			bexists = eval(document.frmcal["selyear"]);
			if(bexists)
			{
				document.frmcal.selyear.style.visibility="hidden";
			}
		}
		// This form is found on recreation/facility_reservation.asp
		var formnames = document.getElementsByName("frmAvail");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmAvail["userid"]);
			if(bexists)
			{
				document.frmAvail.userid.style.visibility="hidden";
			}
			bexists = eval(document.frmAvail["selfacility"]);
			if(bexists)
			{
				document.frmAvail.selfacility.style.visibility="hidden";
			}
		}
		// This form is found on recreation/facility_reporting.asp
		var formnames = document.getElementsByName("frmdate");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmdate["sm"]);
			if(bexists)
			{
				document.frmdate.sm.style.visibility="hidden";
			}
			bexists = eval(document.frmdate["sy"]);
			if(bexists)
			{
				document.frmdate.sy.style.visibility="hidden";
			}
			bexists = eval(document.frmdate["em"]);
			if(bexists)
			{
				document.frmdate.em.style.visibility="hidden";
			}
			bexists = eval(document.frmdate["ey"]);
			if(bexists)
			{
				document.frmdate.ey.style.visibility="hidden";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform0");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform0["weekday"]);
			if(bexists)
			{
				document.availform0.weekday.style.visibility="hidden";
			}
			bexists = eval(document.availform0["beginampm"]);
			if(bexists)
			{
				document.availform0.beginampm.style.visibility="hidden";
			}
			bexists = eval(document.availform0["endampm"]);
			if(bexists)
			{
				document.availform0.endampm.style.visibility="hidden";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform1");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform1["weekday"]);
			if(bexists)
			{
				document.availform1.weekday.style.visibility="hidden";
			}
			bexists = eval(document.availform1["beginampm"]);
			if(bexists)
			{
				document.availform1.beginampm.style.visibility="hidden";
			}
			bexists = eval(document.availform1["endampm"]);
			if(bexists)
			{
				document.availform1.endampm.style.visibility="hidden";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform2");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform2["weekday"]);
			if(bexists)
			{
				document.availform2.weekday.style.visibility="hidden";
			}
			bexists = eval(document.availform2["beginampm"]);
			if(bexists)
			{
				document.availform2.beginampm.style.visibility="hidden";
			}
			bexists = eval(document.availform2["endampm"]);
			if(bexists)
			{
				document.availform2.endampm.style.visibility="hidden";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform3");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform3["weekday"]);
			if(bexists)
			{
				document.availform3.weekday.style.visibility="hidden";
			}
			bexists = eval(document.availform3["beginampm"]);
			if(bexists)
			{
				document.availform3.beginampm.style.visibility="hidden";
			}
			bexists = eval(document.availform3["endampm"]);
			if(bexists)
			{
				document.availform3.endampm.style.visibility="hidden";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform4");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform4["weekday"]);
			if(bexists)
			{
				document.availform4.weekday.style.visibility="hidden";
			}
			bexists = eval(document.availform4["beginampm"]);
			if(bexists)
			{
				document.availform4.beginampm.style.visibility="hidden";
			}
			bexists = eval(document.availform4["endampm"]);
			if(bexists)
			{
				document.availform4.endampm.style.visibility="hidden";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform5");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform5["weekday"]);
			if(bexists)
			{
				document.availform5.weekday.style.visibility="hidden";
			}
			bexists = eval(document.availform5["beginampm"]);
			if(bexists)
			{
				document.availform5.beginampm.style.visibility="hidden";
			}
			bexists = eval(document.availform5["endampm"]);
			if(bexists)
			{
				document.availform5.endampm.style.visibility="hidden";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform6");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform6["weekday"]);
			if(bexists)
			{
				document.availform6.weekday.style.visibility="hidden";
			}
			bexists = eval(document.availform6["beginampm"]);
			if(bexists)
			{
				document.availform6.beginampm.style.visibility="hidden";
			}
			bexists = eval(document.availform6["endampm"]);
			if(bexists)
			{
				document.availform6.endampm.style.visibility="hidden";
			}
		}
		// This form is found on classes/dl_sendmail.asp
		var formnames = document.getElementsByName("frmlocation");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmlocation["SendList"]);
			if(bexists)
			{
				document.frmlocation.SendList.style.visibility="hidden";
			}
			bexists = eval(document.frmlocation["iEmailFormat"]);
			if(bexists)
			{
				document.frmlocation.iEmailFormat.style.visibility="hidden";
			}
		}
		// This form is found on faq/new_faq.asp. faq/manage_faq.asp
		var formnames = document.getElementsByName("NewEvent");
		if (formnames.length == 1)
		{
			var bexists = eval(document.NewEvent["FAQCategoryId"]);
			if(bexists)
			{
				document.NewEvent.FAQCategoryId.style.visibility="hidden";
			}
		}
		// This form is found on classes/roster_list.asp
		var formnames = document.getElementsByName("frmfilter");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmfilter["categoryid"]);
			if(bexists)
			{
				document.frmfilter.categoryid.style.visibility="hidden";
			}
			bexists = eval(document.frmfilter["selDateType"]);
			if(bexists)
			{
				document.frmfilter.selDateType.style.visibility="hidden";
			}
		}
		// This form is found on classes/class_list.asp
		var formnames = document.getElementsByName("ClassForm");
		if (formnames.length == 1)
		{
			var bexists = eval(document.ClassForm["statusid"]);
			if(bexists)
			{
				document.ClassForm.statusid.style.visibility="hidden";
			}
			bexists = eval(document.ClassForm["classtypeid"]);
			if(bexists)
			{
				document.ClassForm.classtypeid.style.visibility="hidden";
			}
			bexists = eval(document.ClassForm["categoryid"]);
			if(bexists)
			{
				document.ClassForm.categoryid.style.visibility="hidden";
			}
			bexists = eval(document.ClassForm["datefilter"]);
			if(bexists)
			{
				document.ClassForm.datefilter.style.visibility="hidden";
			}
		}
		// This form is found on classes/discount_edit.asp
		var formnames = document.getElementsByName("frmdiscount");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmdiscount["discounttypeid"]);
			if(bexists)
			{
				document.frmdiscount.discounttypeid.style.visibility="hidden";
			}
		}
		// This form is found on classes/class_waiver_edit.asp
		var formnames = document.getElementsByName("frmwaiver");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmwaiver["sType"]);
			if(bexists)
			{
				document.frmwaiver.sType.style.visibility="hidden";
			}
		}
		// This form is found on events/updateevent.asp
		var formnames = document.getElementsByName("UpdateEvent");
		if (formnames.length == 1)
		{
			var bexists = eval(document.UpdateEvent["Hour"]);
			if(bexists)
			{
				document.UpdateEvent.Hour.style.visibility="hidden";
			}
			bexists = eval(document.UpdateEvent["Minute"]);
			if(bexists)
			{
				document.UpdateEvent.Minute.style.visibility="hidden";
			}
			bexists = eval(document.UpdateEvent["AMPM"]);
			if(bexists)
			{
				document.UpdateEvent.AMPM.style.visibility="hidden";
			}
			bexists = eval(document.UpdateEvent["DurationInterval"]);
			if(bexists)
			{
				document.UpdateEvent.DurationInterval.style.visibility="hidden";
			}
			bexists = eval(document.UpdateEvent["Category"]);
			if(bexists)
			{
				document.UpdateEvent.Category.style.visibility="hidden";
			}
			bexists = eval(document.UpdateEvent["Timezone"]);
			if(bexists)
			{
				document.UpdateEvent.Timezone.style.visibility="hidden";
			}
		}
		// This form is found on events/recurevent.asp
		var formnames = document.getElementsByName("RecurEvent");
		if (formnames.length == 1)
		{
			var bexists = eval(document.RecurEvent["Recur"]);
			if(bexists)
			{
				document.RecurEvent.Recur.style.visibility="hidden";
			}
			bexists = eval(document.RecurEvent["Ordinal"]);
			if(bexists)
			{
				document.RecurEvent.Ordinal.style.visibility="hidden";
			}
			bexists = eval(document.RecurEvent["DayLike"]);
			if(bexists)
			{
				document.RecurEvent.DayLike.style.visibility="hidden";
			}
			bexists = eval(document.RecurEvent["Month"]);
			if(bexists)
			{
				document.RecurEvent.Month.style.visibility="hidden";
			}
			bexists = eval(document.RecurEvent["DayPick"]);
			if(bexists)
			{
				document.RecurEvent.DayPick.style.visibility="hidden";
			}
			bexists = eval(document.RecurEvent["MonthPick"]);
			if(bexists)
			{
				document.RecurEvent.MonthPick.style.visibility="hidden";
			}
		}
		// This form is found on events/newevent.asp
		var formnames = document.getElementsByName("NewEvent");
		if (formnames.length == 1)
		{
			var bexists = eval(document.NewEvent["Hour"]);
			if(bexists)
			{
				document.NewEvent.Hour.style.visibility="hidden";
			}
			bexists = eval(document.NewEvent["Minute"]);
			if(bexists)
			{
				document.NewEvent.Minute.style.visibility="hidden";
			}
			bexists = eval(document.NewEvent["AMPM"]);
			if(bexists)
			{
				document.NewEvent.AMPM.style.visibility="hidden";
			}
			bexists = eval(document.NewEvent["DurationInterval"]);
			if(bexists)
			{
				document.NewEvent.DurationInterval.style.visibility="hidden";
			}
			bexists = eval(document.NewEvent["Category"]);
			if(bexists)
			{
				document.NewEvent.Category.style.visibility="hidden";
			}
			bexists = eval(document.NewEvent["Timezone"]);
			if(bexists)
			{
				document.NewEvent.Timezone.style.visibility="hidden";
			}
		}
		// This form is found on payments/action_line_list.asp
		var formnames = document.getElementsByName("form1");
		if (formnames.length == 1)
		{
			var bexists = eval(document.form1["orderBy"]);
			if(bexists)
			{
				document.form1.orderBy.style.visibility="hidden";
			}
		}
}

function mitem_unhidethings()
{
		// This form is found on security/edit_user_security.asp and security/copy_user_security.asp
		var formnames = document.getElementsByName("UserForm");
		if (formnames.length == 1)
		{
			var bexists = eval(document.UserForm["iUserID"]);
			if(bexists)
			{
				document.UserForm.iUserID.style.visibility="visible";
			}
			bexists = eval(document.UserForm["iFromUserID"]);
			if(bexists)
			{
				document.UserForm.iFromUserID.style.visibility="visible";
			}
			bexists = eval(document.UserForm["iToUserID"]);
			if(bexists)
			{
				document.UserForm.iToUserID.style.visibility="visible";
			}
		}
		// This form is found on admin/manage_features.asp
		var formnames = document.getElementsByName("pickForm");
		if (formnames.length == 1)
		{
			var bexists = eval(document.pickForm["orgid"]);
			if(bexists)
			{
				document.pickForm.orgid.style.visibility="visible";
			}
		}
		// This form is found on poolpass/poolpass_form.asp
		var formnames = document.getElementsByName("BuyerForm");
		if (formnames.length == 1)
		{
			var bexists = eval(document.BuyerForm["userid"]);
			if(bexists)
			{
				document.BuyerForm.userid.style.visibility="visible";
			}
		}
		// This form is found on poolpass/poolpass_rates.asp
		var formnames = document.getElementsByName("rateform0");
		if (formnames.length == 1)
		{
			var bexists = eval(document.rateform0["iPeriodId"]);
			if(bexists)
			{
				document.rateform0.iPeriodId.style.visibility="visible";
			}
		}
		// This form is found on poolpass/poolpass_rates.asp
		var formnames = document.getElementsByName("rateform1");
		if (formnames.length == 1)
		{
			var bexists = eval(document.rateform1["iPeriodId"]);
			if(bexists)
			{
				document.rateform1.iPeriodId.style.visibility="visible";
			}
		}
		// This form is found on poolpass/poolpass_type_report.asp, classes/class_statisticsreport.asp
		var formnames = document.getElementsByName("YearForm");
		if (formnames.length == 1)
		{
			var bexists = eval(document.YearForm["iyear"]);
			if(bexists)
			{
				document.YearForm.iyear.style.visibility="visible";
			}
		}
		// This form is found on gifts/gift_form.asp
		var formnames = document.getElementsByName("frmpayment");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmpayment["gift"]);
			if(bexists)
			{
				document.frmpayment.gift.style.visibility="visible";
			}
			bexists = eval(document.frmpayment["userid"]);
			if(bexists)
			{
				document.frmpayment.userid.style.visibility="visible";
			}
		}
		// This form is found on recreation/facility_calendar.asp
		var formnames = document.getElementsByName("frmcal");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmcal["selfacility"]);
			if(bexists)
			{
				document.frmcal.selfacility.style.visibility="visible";
			}
			bexists = eval(document.frmcal["selmonth"]);
			if(bexists)
			{
				document.frmcal.selmonth.style.visibility="visible";
			}
			bexists = eval(document.frmcal["selyear"]);
			if(bexists)
			{
				document.frmcal.selyear.style.visibility="visible";
			}
		}
		// This form is found on recreation/facility_reservation.asp
		var formnames = document.getElementsByName("frmAvail");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmAvail["userid"]);
			if(bexists)
			{
				document.frmAvail.userid.style.visibility="visible";
			}
			bexists = eval(document.frmAvail["selfacility"]);
			if(bexists)
			{
				document.frmAvail.selfacility.style.visibility="visible";
			}
		}
		// This form is found on recreation/facility_reporting.asp
		var formnames = document.getElementsByName("frmdate");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmdate["sm"]);
			if(bexists)
			{
				document.frmdate.sm.style.visibility="visible";
			}
			bexists = eval(document.frmdate["sy"]);
			if(bexists)
			{
				document.frmdate.sy.style.visibility="visible";
			}
			bexists = eval(document.frmdate["em"]);
			if(bexists)
			{
				document.frmdate.em.style.visibility="visible";
			}
			bexists = eval(document.frmdate["ey"]);
			if(bexists)
			{
				document.frmdate.ey.style.visibility="visible";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform0");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform0["weekday"]);
			if(bexists)
			{
				document.availform0.weekday.style.visibility="visible";
			}
			bexists = eval(document.availform0["beginampm"]);
			if(bexists)
			{
				document.availform0.beginampm.style.visibility="visible";
			}
			bexists = eval(document.availform0["endampm"]);
			if(bexists)
			{
				document.availform0.endampm.style.visibility="visible";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform1");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform1["weekday"]);
			if(bexists)
			{
				document.availform1.weekday.style.visibility="visible";
			}
			bexists = eval(document.availform1["beginampm"]);
			if(bexists)
			{
				document.availform1.beginampm.style.visibility="visible";
			}
			bexists = eval(document.availform1["endampm"]);
			if(bexists)
			{
				document.availform1.endampm.style.visibility="visible";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform2");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform2["weekday"]);
			if(bexists)
			{
				document.availform2.weekday.style.visibility="visible";
			}
			bexists = eval(document.availform2["beginampm"]);
			if(bexists)
			{
				document.availform2.beginampm.style.visibility="visible";
			}
			bexists = eval(document.availform2["endampm"]);
			if(bexists)
			{
				document.availform2.endampm.style.visibility="visible";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform3");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform3["weekday"]);
			if(bexists)
			{
				document.availform3.weekday.style.visibility="visible";
			}
			bexists = eval(document.availform3["beginampm"]);
			if(bexists)
			{
				document.availform3.beginampm.style.visibility="visible";
			}
			bexists = eval(document.availform3["endampm"]);
			if(bexists)
			{
				document.availform3.endampm.style.visibility="visible";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform4");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform4["weekday"]);
			if(bexists)
			{
				document.availform4.weekday.style.visibility="visible";
			}
			bexists = eval(document.availform4["beginampm"]);
			if(bexists)
			{
				document.availform4.beginampm.style.visibility="visible";
			}
			bexists = eval(document.availform4["endampm"]);
			if(bexists)
			{
				document.availform4.endampm.style.visibility="visible";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform5");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform5["weekday"]);
			if(bexists)
			{
				document.availform5.weekday.style.visibility="visible";
			}
			bexists = eval(document.availform5["beginampm"]);
			if(bexists)
			{
				document.availform5.beginampm.style.visibility="visible";
			}
			bexists = eval(document.availform5["endampm"]);
			if(bexists)
			{
				document.availform5.endampm.style.visibility="visible";
			}
		}
		// This form is found on recreation/facility_availability.asp
		var formnames = document.getElementsByName("availform6");
		if (formnames.length == 1)
		{
			var bexists = eval(document.availform6["weekday"]);
			if(bexists)
			{
				document.availform6.weekday.style.visibility="visible";
			}
			bexists = eval(document.availform6["beginampm"]);
			if(bexists)
			{
				document.availform6.beginampm.style.visibility="visible";
			}
			bexists = eval(document.availform6["endampm"]);
			if(bexists)
			{
				document.availform6.endampm.style.visibility="visible";
			}
		}
		// This form is found on classes/dl_sendmail.asp
		var formnames = document.getElementsByName("frmlocation");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmlocation["SendList"]);
			if(bexists)
			{
				document.frmlocation.SendList.style.visibility="visible";
			}
			bexists = eval(document.frmlocation["iEmailFormat"]);
			if(bexists)
			{
				document.frmlocation.iEmailFormat.style.visibility="visible";
			}
		}
		// This form is found on faq/new_faq.asp, faq/manage_faq.asp
		var formnames = document.getElementsByName("NewEvent");
		if (formnames.length == 1)
		{
			var bexists = eval(document.NewEvent["FAQCategoryId"]);
			if(bexists)
			{
				document.NewEvent.FAQCategoryId.style.visibility="visible";
			}
		}
		// This form is found on classes/roster_list.asp
		var formnames = document.getElementsByName("frmfilter");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmfilter["categoryid"]);
			if(bexists)
			{
				document.frmfilter.categoryid.style.visibility="visible";
			}
			bexists = eval(document.frmfilter["selDateType"]);
			if(bexists)
			{
				document.frmfilter.selDateType.style.visibility="visible";
			}
		}
		// This form is found on classes/class_list.asp
		var formnames = document.getElementsByName("ClassForm");
		if (formnames.length == 1)
		{
			var bexists = eval(document.ClassForm["statusid"]);
			if(bexists)
			{
				document.ClassForm.statusid.style.visibility="visible";
			}
			bexists = eval(document.ClassForm["classtypeid"]);
			if(bexists)
			{
				document.ClassForm.classtypeid.style.visibility="visible";
			}
			bexists = eval(document.ClassForm["categoryid"]);
			if(bexists)
			{
				document.ClassForm.categoryid.style.visibility="visible";
			}
			bexists = eval(document.ClassForm["datefilter"]);
			if(bexists)
			{
				document.ClassForm.datefilter.style.visibility="visible";
			}
		}
		// This form is found on classes/discount_edit.asp
		var formnames = document.getElementsByName("frmdiscount");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmdiscount["discounttypeid"]);
			if(bexists)
			{
				document.frmdiscount.discounttypeid.style.visibility="visible";
			}
		}
		// This form is found on classes/class_waiver_edit.asp
		var formnames = document.getElementsByName("frmwaiver");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmwaiver["sType"]);
			if(bexists)
			{
				document.frmwaiver.sType.style.visibility="visible";
			}
		}
		// This form is found on events/updateevent.asp
		var formnames = document.getElementsByName("UpdateEvent");
		if (formnames.length == 1)
		{
			var bexists = eval(document.UpdateEvent["Hour"]);
			if(bexists)
			{
				document.UpdateEvent.Hour.style.visibility="visible";
			}
			bexists = eval(document.UpdateEvent["Minute"]);
			if(bexists)
			{
				document.UpdateEvent.Minute.style.visibility="visible";
			}
			bexists = eval(document.UpdateEvent["AMPM"]);
			if(bexists)
			{
				document.UpdateEvent.AMPM.style.visibility="visible";
			}
			bexists = eval(document.UpdateEvent["DurationInterval"]);
			if(bexists)
			{
				document.UpdateEvent.DurationInterval.style.visibility="visible";
			}
			bexists = eval(document.UpdateEvent["Category"]);
			if(bexists)
			{
				document.UpdateEvent.Category.style.visibility="visible";
			}
			bexists = eval(document.UpdateEvent["Timezone"]);
			if(bexists)
			{
				document.UpdateEvent.Timezone.style.visibility="visible";
			}
		}
		// This form is found on events/recurevent.asp
		var formnames = document.getElementsByName("RecurEvent");
		if (formnames.length == 1)
		{
			var bexists = eval(document.RecurEvent["Recur"]);
			if(bexists)
			{
				document.RecurEvent.Recur.style.visibility="visible";
			}
			bexists = eval(document.RecurEvent["Ordinal"]);
			if(bexists)
			{
				document.RecurEvent.Ordinal.style.visibility="visible";
			}
			bexists = eval(document.RecurEvent["DayLike"]);
			if(bexists)
			{
				document.RecurEvent.DayLike.style.visibility="visible";
			}
			bexists = eval(document.RecurEvent["Month"]);
			if(bexists)
			{
				document.RecurEvent.Month.style.visibility="visible";
			}
			bexists = eval(document.RecurEvent["DayPick"]);
			if(bexists)
			{
				document.RecurEvent.DayPick.style.visibility="visible";
			}
			bexists = eval(document.RecurEvent["MonthPick"]);
			if(bexists)
			{
				document.RecurEvent.MonthPick.style.visibility="visible";
			}
		}
		// This form is found on events/newevent.asp
		var formnames = document.getElementsByName("NewEvent");
		if (formnames.length == 1)
		{
			var bexists = eval(document.NewEvent["Hour"]);
			if(bexists)
			{
				document.NewEvent.Hour.style.visibility="visible";
			}
			bexists = eval(document.NewEvent["Minute"]);
			if(bexists)
			{
				document.NewEvent.Minute.style.visibility="visible";
			}
			bexists = eval(document.NewEvent["AMPM"]);
			if(bexists)
			{
				document.NewEvent.AMPM.style.visibility="visible";
			}
			bexists = eval(document.NewEvent["DurationInterval"]);
			if(bexists)
			{
				document.NewEvent.DurationInterval.style.visibility="visible";
			}
			bexists = eval(document.NewEvent["Category"]);
			if(bexists)
			{
				document.NewEvent.Category.style.visibility="visible";
			}
			bexists = eval(document.NewEvent["Timezone"]);
			if(bexists)
			{
				document.NewEvent.Timezone.style.visibility="visible";
			}
		}
		// This form is found on payments/action_line_list.asp
		var formnames = document.getElementsByName("form1");
		if (formnames.length == 1)
		{
			var bexists = eval(document.form1["orderBy"]);
			if(bexists)
			{
				document.form1.orderBy.style.visibility="visible";
			}
		}
}



// --------------------------------------------------------------------------------
// that's all folks
