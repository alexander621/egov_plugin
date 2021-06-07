/*
  --- menu items --- 
  note that this structure has changed its format since previous version.
  additional third parameter is added for item scope settings.
  Now this structure is compatible with Tigra Menu GOLD.
  Format description can be found in product documentation.
*/
var MENU_ITEMS = [
	['Home', '../../default.asp', null,
		['Calendar', '../../events/default.asp', null,
			['New Events','../../events/newevent.asp'],
			['View\Manage Event Categories','../../events/eventcategories.asp']
		],
		['Documents', 'http://www.softcomplex.com/support.html'],
		['Security', 'http://www.softcomplex.com/support.html'],
		['Action Line', 'http://www.softcomplex.com/support.html'],
		['Payments', 'http://www.softcomplex.com/support.html'],
		['New Requests', 'http://www.softcomplex.com/support.html'],
		['Registration', 'http://www.softcomplex.com/support.html'],
		['Recreation', '../../recreation/default.asp', null,
			['Facility Management'],
			['Classes\Events'],
			['Commemorative Gifts'],
			['Pool Passes']
		],
		['Logout', '../../signoff.asp', null,]
	]
,];

