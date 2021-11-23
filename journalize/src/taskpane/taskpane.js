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

        Office.context.mailbox.getCallbackTokenAsync(function(result) {
          debugger;

          var token = result.value;
          var ewsurl = Office.context.mailbox.ewsUrl;
          var itemId = Office.context.mailbox.item.itemId;
          var envelope = getSoapEnvelope(itemId);

          var xhttp = new XMLHttpRequest();
          xhttp.withCredentials = true;
          xhttp.open("POST", ewsurl, true);
          xhttp.setRequestHeader("Accept", "application/json");
          xhttp.setRequestHeader("Content-type", "application/xml");
          xhttp.setRequestHeader("Authorization", "Bearer " + token);
          xhttp.responseType = 'json';
          xhttp.send(envelope);

          xhttp.onload = function() {
            if (xhr.status != 200) { // analyze HTTP status of the response
              sendMemoError("error happened, try again or contact it@metz.dk");
            } else { // show the result
              sendMemoSuccess(result);
            }
          };

          xhttp.onprogress = function(event) {
            sendMemoProgress(event);
          };

          xhttp.onerror = function() { // only triggers if the request couldn't be made at all
            sendMemoError("Request failed (probably CORS)");
          };
        });

        function sendMemoSuccess(result) {
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
        }

        function sendMemoError(txt) {
          searchEl.empty();
          $("<p>")
          .addClass("color-red")
          .text(txt).appendTo(searchEl);
        }

        function sendMemoProgress(event) {
          if (event.lengthComputable) {
            var txt = `Received ${event.loaded} of ${event.total} bytes`;
          } else {
            var txt = `Received ${event.loaded} bytes`; // no Content-Length
          }

          searchEl.empty();
          $("<p>")
          .addClass("color-red")
          .text(txt).appendTo(searchEl);
        }
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

  function getSoapEnvelope(itemId) {
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

    '  <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '    <ItemShape>' +
    '      <t:BaseShape>IdOnly</t:BaseShape>' +
    '      <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
    '      <AdditionalProperties xmlns="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '        <FieldURI FieldURI="item:Subject" />' +
    '      </AdditionalProperties>' +
    '    </ItemShape>' +
    '    <ItemIds>' +
    '      <t:ItemId Id="' + itemId + '" />' +
    '    </ItemIds>' +
    '  </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

    return result;
  }

})();