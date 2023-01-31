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
      $("form[name='search'] input").focus();

      // When user press on search button
      $("form[name='search']").on('submit', function(e){
        e.preventDefault();

        let action = document.querySelector('#action').value;
        let keyword = document.getElementById("keyword").value;
        var requestUrl = 'https://api.metz.dk/journalize/v1/search?action=' + action + '&keyword=' + encodeURI(keyword);
        var outputEl = $(".search-result").empty();

        outputEl.append('<p class="color-blue">... please wait...</p>');

        var xhttp = new XMLHttpRequest();
        xhttp.open("GET", requestUrl, true);
        xhttp.send();
        
        xhttp.onload = function() {
          if (xhttp.status != 200) { // analyze HTTP status of the response
            printError(outputEl);
          } else { // show the result
            var res  = JSON.parse(this.responseText);
            if (res.status===1) {
              buildSearchResult(outputEl, res);
            }
            else {
              printError(outputEl, res.message);
            }
          }
        };

        xhttp.onerror = function() { // only triggers if the request couldn't be made at all
          printError(outputEl);
        };

        function buildSearchResult(parent, data) {
          parent.empty();
          let action = data.action;
          let docs = data.docs;
      
          if (docs.length > 0) {
            $("<p>").addClass("color-green").text(docs.length+" document(s) displayed (total: "+data.total+")").appendTo(parent);
            let list = $("<ul>").addClass("my-3").appendTo(parent);
            for (var i = 0; i < docs.length; i++) {
              let li = $("<li>").appendTo(list);
              $("<input>")
              .attr('type', 'radio')
              .attr('name', 'doc')
              .attr('id', "doc"+docs[i].unid)
              .attr('required', 'required')
              .val(action+"|"+docs[i].unid)
              .appendTo(li);
              li.append('<label class="ml-1" for="doc'+docs[i].unid+'">'+docs[i].title+'</label>');
            }
          }
          else {
            $("<p>").addClass("color-green").text("No documents found").appendTo(parent);
          }

          $("<button>")
          .attr("type", "submit")
          .text("Journalize").appendTo(parent);
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
        var outputEl = $(".search-result");

        outputEl.html("<p class='color-blue'>... sending data (please wait) ...</p>");

        Office.context.mailbox.getCallbackTokenAsync(function(result) {
          if (result.status !== "succeeded") {
            printError(outputEl, "Error happened (accesss token was not issued), try again or contact it@metz.dk");
            return;
          }
          
          const token = result.value;
          const ewsurl = Office.context.mailbox.restUrl;
          const ewsItemId = Office.context.mailbox.item.itemId;
          const itemId = Office.context.mailbox.convertToRestId(ewsItemId, Office.MailboxEnums.RestVersion.v2_0);
          const isFromSharedFolder = Office.context.mailbox.initialData.isFromSharedFolder;

          // shared folder
          if (isFromSharedFolder) {
            Office.context.mailbox.item.getSharedPropertiesAsync(function(result) {
              const user = result.value.targetMailbox; 
              linkMemo(token, itemId, ewsurl, docid, user, app, outputEl);
            });
          }
          // private email
          else {
            const user = "me"; 
            linkMemo(token, itemId, ewsurl, docid, user, app, outputEl);
          }
        });
      });
    });

    function linkMemo(token, itemId, ewsurl, docid, user, app, outputEl) {
      const json = {
        "token": token,
        "itemid": itemId,
        "ewsurl": ewsurl,
        "docid": docid,
        "user": user
      };

      var endpoint = "https://api.metz.dk/journalize/v1/" + app;
      var xhttp = new XMLHttpRequest();
      xhttp.open("POST", endpoint, true);
      xhttp.setRequestHeader("Content-type", "application/json");
      xhttp.send(JSON.stringify(json));

      xhttp.onload = function() {
        if (xhttp.status != 200) { // analyze HTTP status of the response
          printError(outputEl);
        } else { // show the result
          var res  = JSON.parse(this.responseText);
          if (res.status===1) {
            confirmLink(outputEl, res);
          }
          else {
            printError(outputEl, res.message);
          }
        }
      };

      xhttp.onerror = function() { // only triggers if the request couldn't be made at all
        printError(outputEl);
      };
    
      function confirmLink(parent, data) {
        var app = $("#app-journalize #action option:selected").text();
        parent.empty();
        parent.append('<p class="color-green">Mail journalized succesfully</p>');
        parent.append('<p><a href="'+data.memo+'" target="_blank">View Notes mail</a></p>');
        parent.append('<p><a href="'+data.doc+'" target="_blank">View '+app+' document</a></p>');
      }
    }

    function printError(el, message) {
      el.empty();
      message = message || "Error happened, try again or contact it@metz.dk";
      $("<p>")
      .addClass("color-red")
      .text(message).appendTo(el);
    }

  };

})();