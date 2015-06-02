/// <reference path="../App.js" />
/// <reference path="../shCore.js"/>
/*global app*/

(function () {
	'use strict';

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();

			$('#get-data-from-selection').click(getDataFromSelection);
		});
	};

	// Reads data from current document selection and displays a notification
	function getDataFromSelection() {
		Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
			function (result) {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					var output = js_beautify(result.value);
					$("#code1").text(output);
                    var codeFather = document.querySelector('.dp-highlighter');
                    
                    if(codeFather != null){
                        codeFather.parentNode.removeChild(codeFather);
                    }
                    
					dp.SyntaxHighlighter.HighlightAll('code', true,false);
                    var htmlContent = SetStyle(document.querySelector('.dp-highlighter'));
                    
                    var content = "";
                    var type = 'html';
                    if(document.querySelector('#withStyle').checked){
                        content = htmlContent;
                        type = 'html';
                    }
                    else{
                        content = output;
                        type = 'text';
                    }
                    if(!document.querySelector('#noreplace').checked){
                        Office.context.document.setSelectedDataAsync(content, {coercionType: type}, function (asyncResult) {
                            app.showNotification(asyncResult.error.message);
                        });
                    }
				} else {
					app.showNotification('Error:', result.error.message);
				}
			}
		);
	}
    
    function SetStyle(codeElem){
        codeElem.querySelector('ol').innerHTML +="<li><span/></li>";
        return document.body.querySelector('#template').innerHTML.replace('%code%',codeElem.outerHTML).replace('<!--','').replace('-->','');
    }
    
})();