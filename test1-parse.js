// this file reads the ical file dumped by the appserver, and builds a little text report of all the events for today.  
// most of the complexity comes from timezones and recurrence relations, which are subtle.

"use strict";

var ical = require('ical');

//var cal = ical.parseFile('test.ics');;
var cal = ical.parseFile('latest.ics');;

//console.log(cal);

var targetDate, targetEnd;;

if (process.argv[2]) {
    targetDate = new Date(process.argv[2]);
    targetDate.setTime(targetDate.getTime() + 1000 * 60 * targetDate.getTimezoneOffset());
} else {
  targetDate = new Date();
  targetDate.setHours(0,0,0,0);
}
targetEnd = new Date();
targetEnd.setTime(targetDate.getTime() + 24*60*60*1000);
console.log("target date: "+targetDate.toISOString().slice(0,10));
//console.log("target end: "+targetEnd.toISOString());


var i; 
var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

var day_events = [];

Object.keys(cal).forEach(function(key) {
//    console.log(cal[key].start.getDate());
    var ev = cal[key];
    //console.log(ev, "start.getDate()", ev.start, " - ", targetDate);
    if (ev.rrule) {
//        console.log(ev.rrule.between(new Date(2015,3,8), new Date(2015,3,15)));
      var instances = ev.rrule.between(targetDate, targetEnd);
      ev.instances = instances;
      if (instances.length > 0) {
        //console.log("Recurrence:", ev.summary, instances);
        day_events.push(ev);
        if (0) {
          console.log("Recurring: ",
                      ev.summary,
                      'is in',
                      ev.location,
                      'on', ev.start.getDate(), months[ev.start.getMonth()], 1900+ev.start.getYear(),
                      ((ev.start && ev.end) ?
                       ('from '+ ('0'+ev.start.getHours()).slice(-2)+('0'+ev.start.getMinutes()).slice(-2)+
                        ' to '+ ('0'+ev.end.getHours()).slice(-2)+('0'+ev.end.getMinutes()).slice(-2)):
                       ''));
        }
        //console.log(ev);
      }
    } else if (ev.start >= targetDate && ev.start < targetEnd) {
      day_events.push(ev);
      
      if (0) {
        console.log("Non-recurring: ",
                    ev.summary,
                    'is in',
                    ev.location,
                    'on', ev.start.getDate(), months[ev.start.getMonth()], 1900+ev.start.getYear(),
                    ((ev.start && ev.end) ?
                     ('from '+ ('0'+ev.start.getHours()).slice(-2)+('0'+ev.start.getMinutes()).slice(-2)+
                      ' to '+ ('0'+ev.end.getHours()).slice(-2)+('0'+ev.end.getMinutes()).slice(-2)):
                     '')); 
      }
    };
});

day_events.forEach(function(ev,i,a) {
  if (ev.hasOwnProperty("instances")) {
//    console.log(ev.instances);
    ev.start = ev.instances[0];
  }
});

var events = day_events.sort(function(a,b) {
  return a.start - b.start;
});

events.forEach(function(ev,i,a) {
  // magic to flag things in outlook not to show here
  if (ev.categories.filter(function(x) { return x === 'Yellow Category'; }).length > 0) { return; }
  console.log(""+
              (((typeof(ev.start) !== 'undefined')  && (typeof(ev.end) !== 'undefined')) ?
               (('0'+ev.start.getHours()).slice(-2)+('0'+ev.start.getMinutes()).slice(-2)+
                ' to '+ ('0'+ev.end.getHours()).slice(-2)+('0'+ev.end.getMinutes()).slice(-2))+":"
               :""),
              ev.summary,
              '@',
              ev.location
  );
  
});

process.exit(0);



