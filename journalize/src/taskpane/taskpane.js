/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

(function(){
  'use strict';

  Office.initialize = function(reason){

    // search
    jQuery(document).ready(function(){

      // When user press on search button
      $("form[name='search'").on('submit', function(e){
        e.preventDefault();

        let app = document.querySelector('input[name="app"]:checked').value;
        let keyword = document.getElementById("keyword").value;
        var requestUrl = 'https://api-dev.metz.dk/journalize/v1/search?app=' + app + '&keyword=' + keyword;
        var searchEl = $(".search-result").empty();

        $("<p>").addClass("color-blue").text("please wait...").appendTo(searchEl);

        var xhttp = new XMLHttpRequest();
        xhttp.open("GET", requestUrl, true);
        xhttp.send();
        
        xhttp.onload = function() {
          searchEl.empty();
          if (xhttp.status != 200) { // analyze HTTP status of the response
            sendMemoError("Error happened, try again or contact it@metz.dk");
          } else { // show the result
            buildSearchResult(searchEl, JSON.parse(this.responseText));
          }
        };

        xhttp.onerror = function() { // only triggers if the request couldn't be made at all
          searchEl.empty();
          $("<p>")
          .addClass("color-red")
          .text("Error happened, try again or contact it@metz.dk").appendTo(searchEl);
        };

        function buildSearchResult(parent, data) {
          let app = data.app;
          let docs = data.docs;
      
          if (docs.length > 0) {
            $("<p>").addClass("color-green").text(`${docs.length} document(s) displayed (total: ${data.total})`).appendTo(parent);
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
              .text(docs[i].title)
              .appendTo(li);
            }
          }
          else {
            $("<p>").addClass("color-green").text("No documents found").appendTo(parent);
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
      });
    });

    // search-result submission
    jQuery(document).ready(function(){
      // When user press on search button
      $("form[name='search-result'").on('submit', function(e){
        e.preventDefault();

        var docid = $("input[type=radio][name='doc']:checked").val();
        var app = docid.split('|')[0];
        var searchEl = $(".search-result");

        searchEl.html("... sending data (please wait) ...");

        // TEST: GRAPH
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
          if (result.status === "succeeded") {
              var accessToken = result.value;
              getCurrentItem(accessToken);
          } else {
              // Handle the error
          }
        });

        function getCurrentItem(accessToken) {
          // Get the item's REST ID
          var itemId = getItemRestId();
      
          // Construct the REST URL to the current item
          // Details for formatting the URL can be found at 
          // https://msdn.microsoft.com/office/office365/APi/mail-rest-operations#get-a-message-rest
         // var getMessageUrl = Office.context.mailbox.restUrl +
         //     '/v2.0/me/messages/' + itemId;'
           var PostURL = "https://outlook.office.com/api/beta/me/GetMailTips";
           var mtipRequest = "{ \"EmailAddresses\": [ \"gscales@datarumble.com\" ],\"MailTipsOptions\": \"automaticReplies, mailboxFullStatus\"}";
          //var getMessageUrl = "https://graph.microsoft.com/beta/me/messages/" + itemId;
          $.ajax({
              type: "POST",
              contentType: "application/json; charset=utf-8",
              url: PostURL,
              data: mtipRequest,
              dataType: 'json',
              headers: { 'Authorization': 'Bearer ' + accessToken }
          }).done(function (item) {
              // Message is passed in `item`
              var subject = item.value[0].EmailAddress.Address;
              $('#ServerName').text("Message Subject : " + subject);
          }).fail(function (error) {
              // Handle error
          });
        }
        
        function getItemRestId() {
            if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
                // itemId is already REST-formatted
                return Office.context.mailbox.item.itemId;
            } else {
                // Convert to an item ID for API v2.0
                return Office.context.mailbox.convertToRestId(
                    Office.context.mailbox.item.itemId,
                    Office.MailboxEnums.RestVersion.v2_0
                );
            }
        }

        Office.context.mailbox.getCallbackTokenAsync(function(result) {
          var token = result.value;
          var ewsurl = Office.context.mailbox.ewsUrl;
          var itemId = Office.context.mailbox.item.itemId;

          var json = {
            "token": token,
            "itemid": itemId,
            "ewsurl": ewsurl,
            "docid": docid
          };

          var endpoint = "https://api-dev.metz.dk/journalize/v1/" + app;
          var xhttp = new XMLHttpRequest();
          xhttp.open("POST", endpoint, true);
          xhttp.setRequestHeader("Content-type", "application/json");
          xhttp.send(JSON.stringify(json));

          xhttp.onload = function() {
            if (xhttp.status != 200) { // analyze HTTP status of the response
              sendMemoError("Error happened, try again or contact it@metz.dk");
            } else { // show the result
              searchEl.empty();
              confirmLink(searchEl, JSON.parse(this.responseText));
            }
          };

          xhttp.onerror = function() { // only triggers if the request couldn't be made at all
            sendMemoError("Request failed");
          };

          function sendMemoError(txt) {
            searchEl.empty();
            $("<p>")
            .addClass("color-red")
            .text(txt).appendTo(searchEl);
          }
        
          function confirmLink(parent, data) {
            $("<p>").addClass("color-green").html(data.message).appendTo(parent);
          }
        
        });
     });
    });
  };

})();