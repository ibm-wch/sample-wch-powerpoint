/*
 * Copyright 2018  IBM Corp.
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 * Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the
 * specific language governing permissions and limitations under the License.
 */


'use strict';

//// The initialize function must be run each time a new page is loaded
Office.initialize = function (reason) {
    $(document).ready(function () {
    });
};

/*
 * Get saved wchAPIURL on load
 */
$(document).ready(function(){
	$('#current-wch-id').attr("value", localStorage.getItem('wchAPIURL'));
});

/*
 * Toggle display of options form 
 */
function wchOptions() {
	
	$('#options-form').toggle();
}

/*
 * Save current value of wchAPIURL
 */
function saveOptions(){
	localStorage.setItem('wchAPIURL', $('#current-wch-id')[0].value);
}


// 1. 'addEventListener' is for standards-compliant web browsers and 'attachEvent' is for IE Browsers
var eventMethod = window.addEventListener ? 'addEventListener' : 'attachEvent';
var eventer = window[eventMethod];
// 2. if 'attachEvent', then we need to select 'onmessage' as the event
// else if 'addEventListener', then we need to select 'message' as the event
var messageEvent = eventMethod === 'attachEvent' ? 'onmessage' : 'message';


/*
 * Launch the wch asset picker using the wchAPIURL
 */
function launchPicker(myhandler) {
	var wchID = $('#current-wch-id')[0].value;
	
	//images only
	var url = 'https://www.digitalexperience.ibm.com/content-picker/picker.html?apiUrl='+wchID+ '&fq=classification:asset&fq=assetType:image';

	$('#pickerDialog').dialog({
        autoOpen: false,
        show: 'fade',
        hide: 'fade',
        modal: false,
        height: window.innerHeight - 30,
        resizable: true,
        minHeight: 500,
        maxWidth: 150,
        width: 348,
        position: { my: 'right center', at: 'right center', of: window },
        open: function() {        	
        	
        	$('#pickerIframe').attr('src', url);
        },
        title: 'Find',
    });

	
    // Listen to message from child iFrame window
    eventer(messageEvent, myhandler, false);

    //open the dialog
    $('#pickerDialog').dialog('open');
}

//handle how the chosen image is displayed on the page
function resultHandler(e) {
    $('#pickerDialog').dialog('close');

    var result = JSON.parse(e.data);
    //construct the resource url
    var wchID = $('#current-wch-id')[0].value;
    var akamaiUrl = wchID.replace('/api','') + result.path;
    getDataUri(akamaiUrl,insertImageFromBase64String);
}

/*
 * Use a canvas to get the Base64 representation of the image
 */
function getDataUri(url, callback) {
    var wchImage = new Image();
    wchImage.crossOrigin="anonymous";
    wchImage.onload = function () {
        var canvas = document.createElement('canvas');
        canvas.width = this.width; 
        canvas.height = this.height;

        canvas.getContext('2d').drawImage(this, 0, 0);

        // Get raw image data
        callback(canvas.toDataURL('image/png').replace(/^data:image\/(png|jpg);base64,/, ''));

    };

    wchImage.src = url;
}

/*
 * Insert the image into Powerpoint from the Base64 String
 */
function insertImageFromBase64String(image) {
	  // Call Office.js to insert the image into the document.
	  Office.context.document.setSelectedDataAsync(image, {
	      coercionType: Office.CoercionType.Image
	  },
	      function (asyncResult) {
	          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
	              showNotification("Error", asyncResult.error.message);
	          }
	      });
	}
