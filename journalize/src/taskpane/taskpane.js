/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

(function(){
  'use strict';

  var config;
  var settingsDialog;

  Office.initialize = function(reason){

    // search
    jQuery(document).ready(function(){

      // When user press on search button
      $("form[name='search'").on('submit', function(e){
        e.preventDefault();

        let app = $("input[type=radio][name='app']:checked").val();
        let keyword = $("#keyword").val();
        var requestUrl = 'https://api-dev.metz.dk/journalize/v1/search?app=' + app + '&keyword=' + keyword;
        var searchEl = $(".search-result").empty();
        $.get(requestUrl, function(data) {
          buildSearchResult(searchEl, data);
        })
        .fail(function() {
          $("<p>")
          .addClass("color-red")
          .text("error happened during search, try again or contact it@metz.dk").appendTo(searchEl);
        })
      });
    });

    // search-result submission
    jQuery(document).ready(function(){
      // When user press on search button
      $("form[name='search-result'").on('submit', function(e){
        e.preventDefault();

        var docid = $("input[type=radio][name='doc']:checked").val();
        var searchEl = $(".search-result");

        searchEl.html("... preparing data ...");

        var request = GetItem();
        var envelope = getSoapEnvelope(request);

        Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){ 
          if (result.status === "succeeded") { 
            var accessToken = result.value; 
        
            // Use the access token. 
            getCurrentItem(accessToken); 
          } else { 
            // Handle the error. 
          } 
        }); 
        
        function getItemRestId() { 
          if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {  
            // itemId is already REST-formatted. 
            return Office.context.mailbox.item.itemId; 
          } else { 
            // Convert to an item ID for API v2.0. 
            return Office.context.mailbox.convertToRestId( 
              Office.context.mailbox.item.itemId, 
              Office.MailboxEnums.RestVersion.v2_0 
            ); 
          } 
        } 
        
        function getCurrentItem(accessToken) { 
          // Get the item's REST ID.  
          var itemId = getItemRestId(); 
        
          // Construct the REST URL to the current item. 
          // Details for formatting the URL can be found at 
          // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages. 
          var getMessageUrl = Office.context.mailbox.restUrl + 
            '/v2.0/me/messages/' + itemId;
        
          $.ajax({ 
            url: getMessageUrl, 
            dataType: 'json',  
            headers: { 'Authorization': 'Bearer ' + accessToken }  
          }).done(function(item){ 
            // Message is passed in `item`. 
            var subject = item.Subject; 
          }).fail(function(error){ 
            // Handle error. 
          }); 
        } 

        Office.context.mailbox.makeEwsRequestAsync(envelope, function(result){
          if (result.status === "failed") {
            searchEl.empty();
            $("<p>")
            .addClass("color-red")
            .text(result.error.message).appendTo(searchEl);
            return;
          }

          var parser = new DOMParser();
          var doc = parser.parseFromString(result.value, "text/xml");
          var values = doc.getElementsByTagName("t:MimeContent");
          var subject = doc.getElementsByTagName("t:Subject");
          console.log(subject[0].textContent)
          
          var requestUrl = 'https://api-dev.metz.dk/journalize/v1/link';

          searchEl.html("... sending data (please wait) ...");
          $.post(requestUrl, {"docid": docid, "subject": subject[0].textContent, "body": values[0].textContent})
          .done(function(data) {
            searchEl.empty();
            confirmLink(searchEl, data);
          })
          .fail(function() {
            searchEl.empty();
            $("<p>")
            .addClass("color-red")
            .text("error happened, try again or contact it@metz.dk").appendTo(searchEl);
          })
        });
     });
    });
  };

  function confirmLink(parent, data) {
    $("<p>").addClass("color-green").html(data.message).appendTo(parent);
  }

  function buildSearchResult(parent, data) {
    let app = data.app;
    let docs = data.docs;

    if (docs.length > 0) {
      $("<p>").addClass("color-green").text(docs.length+" document(s) found: ").appendTo(parent);
      let list = $("<ul>").addClass("my-3").appendTo(parent);
      for (var i = 0; i < docs.length; i++) {
        let li = $("<li>").appendTo(list);
        $("<input>")
        .attr('type', 'radio')
        .attr('name', 'doc')
        .attr('id', 'doc' + docs[i].unid)
        .attr('required', 'required')
        .val(`${app}|${docs[i].unid}`)
        .appendTo(li);
        $("<label>")
        .attr('for', 'doc' + docs[i].unid)
        .text(docs[i].form)
        .appendTo(li);
      }
    }
    else {
      $("<p>").addClass("color-green").text("no documents found").appendTo(parent);
    }

    $("<button>")
    .attr("type", "submit")
    .text("Journalize").appendTo(parent);

    let debug = $("<div>").addClass("mt-2").appendTo(parent);
    $("<small>").text("server: " + data.server).appendTo(debug);
    $("<br/>").appendTo(debug);
    $("<small>").text("keyword: " + data.keyword).appendTo(debug);
    $("<br/>").appendTo(debug);
    $("<small>").text("app: " + data.app).appendTo(debug);
  }

  // https://gscales.github.io/OWAExportAsEML/MessageRead.js
  function getSoapEnvelope(request) {
    // Wrap an Exchange Web Services request in a SOAP envelope.
    var result =

    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +

    request +

    '  </soap:Body>' +
    '</soap:Envelope>';

    return result;
  }

  function GetItem() {
      var results =
    '  <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '    <ItemShape>' +
    '      <t:BaseShape>IdOnly</t:BaseShape>' +
    '      <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
    '      <AdditionalProperties xmlns="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '        <FieldURI FieldURI="item:Subject" />' +
    '      </AdditionalProperties>' +
    '    </ItemShape>' +
    '    <ItemIds>' +
    '      <t:ItemId Id="' + Office.context.mailbox.item.itemId + '" />' +
    '    </ItemIds>' +
    '  </GetItem>';
  
      return results;
  }

})();