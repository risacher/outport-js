// Copyright 2006 - Ryan Watkins (ryan@ryanwatkins.net)
//
// version 1.0.4 (29 Oct 2006)
// 
// much code converted to javascript from outlook2ical v1.03
// by Norman L. Jones, Provo Utah (njones61@gmail.com)
// additional fixes by
// Andrew Johnson, Alastair Rankine, Dane Walther, Zan Hecht, Markus Untera

// configuration --------------------------------------------------------

var categories = new Array("Private","Holiday"); // calendar categories not/only to export
var exportMode = "all";                          // "all"  - to export ALL entries (regardless of categories above)
                                                 // "not"  - to export all BUT categories listed; 
                                                 // "only" - to export ONLY those categories listed; 

var includeHistory = 90;          // how many days back to include old events
var icsFilename = "H:\\outport\calendar.ics";  // where to store the file

// Outlook Redemption works around limitations imposed by the Outlook
// Security Patch and Service Pack 2 of MS Office 98/2000 and Office
// 2002 and 2003 (which include Security Patch).
// 
// In our case, the 'body' object for the calendar items - the
// 'description' in the ics files is protected by default and Outlook
// will prompt you to allow access when the script is run.  If you
// install Redemption, available at
//
//   http://www.dimastr.com/redemption/
//
// then you can avoid this issue and automate running this script.
//
// If you cant install Redemption and want to include the
// body/description for the calendar items, despite the promoting from
// Outlook, set includeBody to true.  If this script finds Redemption
// is installed, it will set includeBody to true. 

var includeBody = true;

// if you wish to include reminders with your ical items, set this to 'true'

var includeAlarm = false;

// ----------------------------------------------------------------------

// Outlook constants - from http://www.winscripter.com/
// OlDaysOfWeek
var olSunday    =  1;
var olMonday    =  2;
var olTuesday   =  4;
var olWednesday =  8;
var olThursday  = 16;
var olFriday    = 32;
var olSaturday  = 64;

// OlDefaultFolders Constants 
var olFolderCalendar = 9;

// OlRecurrenceType Constants 
var olRecursDaily    = 0;
var olRecursWeekly   = 1;
var olRecursMonthly  = 2;
var olRecursMonthNth = 3;
var olRecursYearly   = 5;
var olRecursYearNth  = 6;

// OlSensitivity Constants
var olNormal       = 0;
var olPersonal     = 1;
var olPrivate      = 2;
var olConfidential = 3;


function createEvent(item) {

    var event = "BEGIN:VEVENT\n";

    if (item.alldayevent == true) {
        event += "DTSTART;VALUE=DATE:" + formatDate(item.start) + "\n";
        if (item.isrecurring == false) {
            event += "DTEND;VALUE=DATE:" + formatDate(item.end) + "\n";
        }
    } else {
        event += "DTSTART;TZID=US-Eastern:" + formatDateTime(item.start) + "\n";
        event += "DTEND;TZID=US-Eastern:" + formatDateTime(item.end) + "\n";
    }

    if (item.isrecurring == true) {
        event += createReoccuringEvent(item);
    }

    event += "LOCATION:" + item.location + "\n";
    event += "SUMMARY:" + item.subject + "\n";
    
    if (item.categories.length < 1) {
        event += "CATEGORIES:(none)\n";
    } else {
        event += "CATEGORIES:" + item.categories + "\n";
    }

    if (item.sensitivity == olNormal) {  
        event += "CLASS:PUBLIC\n";
    } else if (item.sensitivity == olPersonal) {
        event += "CLASS:CONFIDENTIAL\n";
    } else {
        event += "CLASS:PRIVATE\n";
    }

    //    if (includeBody) {
    //        event += "DESCRIPTION:" + cleanLineEndings(item.body) + "\n";
    //    }
    if (includeAlarm) {
		if (item.reminderminutesbeforestart > 0){
	        event += "BEGIN:VALARM\n";
	        event += "TRIGGER:-PT" + item.reminderminutesbeforestart + "M\n";
	        event += "ACTION:DISPLAY\nDESCRIPTION:Reminder\nEND:VALARM\n";
		}
    }
    event += "END:VEVENT\n";

    return event;
}

function createReoccuringEvent(item) {

    var recurEvent = "RRULE:";

    var pattern = item.getrecurrencepattern;
    var patternType = pattern.recurrencetype;

    if (patternType === olRecursDaily) {

        recurEvent += "FREQ=DAILY";
        if (pattern.noenddate !== true) {
            recurEvent += ";UNTIL=" + formatUTCDateTime(pattern.patternenddate);
            // The end date/time is marked as 12:00am on the last day.
            // When this is parsed by php-ical, the last day of the
            // sequence is missed. The MS Outlook code has the same
            // bug/issue.  To fix this, change the end time from 12:00 am
            // to 11:59:59 pm.
            recurEvent = recurEvent.replace(/T000000?/g,"T235959");
        }
        recurEvent += ";INTERVAL=" + pattern.interval;
    
    } else if (patternType === olRecursMonthly) {

        recurEvent += "FREQ=MONTHLY";
        if (pattern.noenddate !== true) {
            recurEvent += ";UNTIL=" + formatUTCDateTime(pattern.patternenddate);
        }
        recurEvent += ";INTERVAL=" + pattern.interval;
        recurEvent += ";BYMONTHDAY=" + pattern.dayofmonth;

    } else if (patternType === olRecursMonthNth) {

        recurEvent += "FREQ=MONTHLY";
        if (pattern.noenddate !== true) {
            recurEvent += ";UNTIL=" + formatUTCDateTime(pattern.patternenddate);
        }

        recurEvent += ";INTERVAL=" + pattern.interval;
        // php-icalendar has a bug for monthly recurring events.  If
        // it is the last day of the month, you can't use the
        // BYDAY=-1SU option, unless you also do the BYMONTH option
        // (which only is useful for yearly events).  However, the
        // BYWEEK option seems to work for the last week of the month
        // (but not for the first week of the month).  Anyway, this
        // exeception seems to work.
        if (pattern.instance === 5) {
            recurEvent += ";BYWEEK=-1;BYDAY=" + daysOfWeek("", pattern);
            } else {
            recurEvent += ";BYDAY=" + daysOfWeek(weekNum(pattern.instance), pattern);
        }

    } else if (patternType === olRecursWeekly) {

        recurEvent += "FREQ=WEEKLY";
        if (pattern.noenddate !== true) {
            recurEvent += ";UNTIL=" + formatUTCDateTime(pattern.patternenddate);
        }
        recurEvent += ";INTERVAL=" + pattern.interval;
        recurEvent += ";BYDAY=" + daysOfWeek("", pattern);

    } else if (patternType === olRecursYearly) {

        recurEvent += "FREQ=YEARLY";
        if (pattern.noenddate !== true) {
            recurEvent += ";UNTIL=" + formatUTCDateTime(pattern.patternenddate);
        }
        recurEvent += ";INTERVAL=1";
	if (daysOfWeek("", pattern)) {
	   recurEvent += ";BYDAY=" + daysOfWeek("", pattern);
        }
    } else if (patternType === olRecursYearNth) {

        recurEvent += "FREQ=YEARLY";
        if (pattern.noenddate !== true) {
            recurEvent += ";UNTIL=" + formatUTCDateTime(pattern.patternenddate);
        }
        recurEvent += ";BYMONTH=" + monthNum(pattern.monthofyear);
        recurEvent += ";BYDAY=" + daysOfWeek(weekNum(pattern.instance), pattern);

    }
    recurEvent += "\n";
	return recurEvent;
}

function formatDate(date) {
    var oDate = new Date(date);
    icaldate = "" + oDate.getFullYear() + padzero((oDate.getMonth() + 1)) + padzero((oDate.getDate()));
    return icaldate;
}

function formatDateTime(date) {
    var oDate = new Date(date);
    icaldate = "" + oDate.getFullYear() + padzero((oDate.getMonth() + 1)) + padzero((oDate.getDate())) +
        "T" + padzero(oDate.getHours()) + padzero(oDate.getMinutes()) + padzero(oDate.getSeconds());
    return icaldate;
}

function formatUTCDateTime(date) {
    var oDate = new Date(date);
    icaldate = "" + oDate.getUTCFullYear() + padzero((oDate.getUTCMonth() + 1)) + padzero((oDate.getUTCDate())) +
        "T" + padzero(oDate.getUTCHours()) + padzero(oDate.getUTCMinutes()) + padzero(oDate.getUTCSeconds()) + "Z";
    return icaldate;
}

function daysOfWeek(week, pattern) {
    var mask = pattern.dayofweekmask;
    var daysOfWeek = "";

    if (mask & olMonday) {
        daysOfWeek = week + "MO";
    }
    if (mask & olTuesday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "TU";
    }
    if (mask & olWednesday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "WE";
    }
    if (mask & olThursday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "TH";
    }
    if (mask & olFriday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "FR";
    }
    if (mask & olSaturday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "SA";
    }
    if (mask & olSunday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "SU";
    }

    return daysOfWeek;
}

function weekNum(week) {
    if (week == 5) { 
        week = "-1"; 
    } else {
        padzero(week);
    }
    return week;
}

function monthNum(month) {
    var month = month + "";  // incase month comes in as a num
    month = month.toLowerCase().substr(0,3);

    var monthNum = 0;

    if (month == "jan") {
        monthNum = 1;
    } else if (month == "feb") {
        monthNum = 2;
    } else if (month == "mar") {
        monthNum = 3;
    } else if (month == "apr") {
        monthNum = 4;
    } else if (month == "may") {
        monthNum = 5;
    } else if (month == "jun") {
        monthNum = 6;
    } else if (month == "jul") {
        monthNum = 7;
    } else if (month == "aug") {
        monthNum = 8;
    } else if (month == "sep") {
        monthNum = 9;
    } else if (month == "oct") {
        monthNum = 10;
    } else if (month == "nov") {
        monthNum = 11;
    } else if (month == "dec") {
        monthNum = 12;
    } else {
        monthNum = month;
    }

    return monthNum;
}

function padzero(string) {
    if (String(string).length < 2) {
        string = "0" + string;
    }
    return string;
}

function cleanLineEndings(string) {
    string = string.replace(/\r/g,'\n');
    string = string.replace(/\n\n/g,'\n');
    string = string.replace(/\n/g,'\\n');
    string = string.replace(/,/g,'\,');
    return string;
}


function alert(msg) {
	if (1) {
		var myMsgBox=new ActiveXObject("wscript.shell");
		myMsgBox.Popup(msg);
	}	
}

function post_ical(ics) {
	var DataToSend = "id=1";
	//var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP.3.0");
	xmlhttp.Open("POST","http://risacher.org/outport/",false);
	xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	xmlhttp.onreadystatechange=function() {	
		if (xmlhttp.readyState==4 && xmlhttp.status==200) {
//			alert(xmlhttp.responseXML.xml);
			//document.getElementById("myDiv").innerHTML=xmlhttp.responseText;
		} else {
			var debug = "readyState: "+ xmlhttp.readyState;
//			if (typeof(xmlhttp.status) != 'undefined') { debug += ("   status: " + xmlhttp.status.toString()) ; }
//			alert(debug);
		}
	}
	xmlhttp.send(ics);
}

function fetch_calendar() {
	var ics = "BEGIN:VCALENDAR\n" +
		"CALSCALE:GREGORIAN\n" +
		"VERSION:2.0\n";
	
	var ol = new ActiveXObject("outlook.application");
	var calendar = ol.getnamespace("mapi").getdefaultfolder(olFolderCalendar).items;


	var today = new Date();
	var total = calendar.Count;
	var exportItem = true;

	for (var i=1; i<=total; i++) {

		try {
			var rItem = calendar(i);
			var item = new ActiveXObject("Redemption.SafeContactItem");
			item.item = rItem;
			includeBody = true;
		} catch(e) {
			var item = calendar(i);
		}

		if ((exportMode == "not") || (exportMode == "only")) {
        
			exportItem = (exportMode=="not") ? true : false; // setup default
        
			for (var j=0; j < categories.length; j++) {
				if (item.categories.indexOf(categories[j]) != -1) {  // item found
					exportItem = (exportMode == "only") ? true : false;
				}
				continue; // no need to go through the rest of the array... 
			}
		} 

		if (exportItem) {
			if (!item.isrecurring) {
				var date = new Date(item.end);
				if (Math.round(((today - date) / (86400000))) > includeHistory) { continue; }
			}
			ics += createEvent(item);
		}
	}

	ics += "END:VCALENDAR\n";
	//var myMsgBox=new ActiveXObject("wscript.shell");
	//myMsgBox.Popup ("Hey, this works"+ics);
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var icsFH = fso.CreateTextFile(icsFilename, true);
	icsFH.Write(ics);
	icsFH.Close();
	return ics;
}
	
(function main_thing() {
	WScript.Sleep(2*60*1000);
	var ics = fetch_calendar();	
	post_ical(ics);
//	alert ("Hey, this works"+ics);
	WScript.Quit();
}());
	
