// Surfplaza user script
// version 0.1 BETA!
// 2005-04-25
// Copyright (c) 2009, myTSelection.blogspot.com
// Released under the GPL license
// http://www.gnu.org/copyleft/gpl.html
//
// --------------------------------------------------------------------
//
// This is a Greasemonkey user script.  To install it, you need
// Greasemonkey 0.3 or later: http://greasemonkey.mozdev.org/
// Then restart Firefox and revisit this script.
// Under Tools, there will be a new menu item to "Install User Script".
// Accept the default configuration and install.
//
// To uninstall, go to Tools/Manage User Scripts,
// select "Surfplaza", and click Uninstall.
//
// --------------------------------------------------------------------
//
// ==UserScript==
// @name          Surfplaza
// @namespace     http://myTselection.blogspot.com
// @description   Little tweaks for Surfplaza.be
// @include       http://*.surfplaza.be/*
// @include       http://*.webtoday.be/*
// ==/UserScript==

//alert('Surfplaza');
var headerBar = document.getElementById('header');
if (headerBar) {
    headerBar.parentNode.removeChild(headerBar);
}

var combellBlock = document.getElementsByClassName('item style_b special', document.getElementById('homepage'));
if (combellBlock) {
//	alert('combellBlock');
	
	for(var i=0,j=combellBlock.length; i<j; i++) {
//    	combellBlock[i].parentNode.removeChild(combellBlock[i]);
		var combellH3 = getElementsByNodeNameWithinNode('H3', combellBlock[i]);
//		alert('combellH3: ' + combellH3[0].firstChild.textContent);
		combellH3[0].firstChild.data ="Weer";
		
		var combellSpecialBlock = document.getElementsByClassName('special', combellBlock[i])[1];
//		alert("special" + combellSpecialBlock.className);
		//combellBlock[i].removeChild(combellSpecialBlock);
		removeChildNodes(combellSpecialBlock);
	   var weatherImage = document.createElement('img');
	   weatherImage.setAttribute('width','199');
	   weatherImage.setAttribute('alt','Weer vandaag / morgen');
	   weatherImage.setAttribute('title','Weer vandaag / morgen');
	   weatherImage.setAttribute('name','weervandaag');
	   weatherImage.setAttribute('onmouseout',"document.weervandaag.src='http://www.kmi.be/meteo/view/nl/211797-Verwachtingen.html?image=today&amp;ext=.png';");
	   weatherImage.setAttribute('onmouseover',"document.weervandaag.src='http://www.kmi.be/meteo/view/nl/211797-Verwachtingen.html?image=tomorrow&amp;ext=.png';");
	   weatherImage.setAttribute('src','http://www.meteo.be/meteo/view/nl/211797-Verwachtingen.html?image=today&ext=.png');
	   combellSpecialBlock.appendChild(weatherImage);
	}
}
 document.title="WebToday";

function getElementsByClassName(classname, node) {
	if(!node) node = document.getElementsByTagName("body")[0];
	var a = [];
	var re = new RegExp('^' + classname + '$');
	var els = node.getElementsByTagName("*");
	for(var i=0,j=els.length; i<j; i++)
		if(re.test(els[i].className))a.push(els[i]);
	return a;
} 

function getElementsByNodeNameWithinNode(nodenamefilter, node) {
	if(!node) node = document.getElementsByTagName("body")[0];
	var a = [];
	var re = new RegExp('^' + nodenamefilter + '$');
	var els = node.getElementsByTagName("*");
	for(var i=0,j=els.length; i<j; i++)
		if(re.test(els[i].nodeName))a.push(els[i]);
	return a;
} 

function removeChildNodes(ctrl)
{
  while (ctrl.childNodes[0])
  {
    ctrl.removeChild(ctrl.childNodes[0]);
  }
}