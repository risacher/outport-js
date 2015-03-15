"use strict";

var ical = require('ical');

var cal = ical.parseFile('output.dat');;

//console.log(cal);

var i; 
var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

Object.keys(cal).forEach(function(key) {
//    console.log(cal[key].start.getDate());
    var ev = cal[key];
    console.log(ev.summary,
            'is in',
            ev.location,
            'on', ev.start.getDate(), months[ev.start.getMonth()]);

});





