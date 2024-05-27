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
      $("form[name='search'] select").focus();

      // When user changes the action selection
      $("form[name='search'] select").on('change', function() {
        let action = $('#action').val();
        let approvalContainer = $('.js-approval');
  
        const flag = action === 'order-open' || action === 'order-closed' || action === 'order-rma';
        approvalContainer.toggle(flag);

        // Reset the approval checkbox
        $('#approval').prop('checked', false);
      });

      // When user press on search button
      $("form[name='search']").on('submit', function(e){
        e.preventDefault();

        let action = document.querySelector('#action').value;
        let keyword = document.getElementById("keyword").value;
        let recentControl = document.getElementById("recent");
        let recent = recentControl.checked ? "1" : "0";

        var requestUrl = 'https://api.metz.dk/journalize/v1/search?action=' + action + '&keyword=' + encodeURI(keyword) + '&recent=' + recent;
        var searchSection = $(".js-search-section").hide();
        var searchResult = $(".js-search-result").empty();
        var searchStatus = $(".js-search-status").empty();
        searchStatus.append('<p class="color-blue">... please wait...</p>');

        var xhttp = new XMLHttpRequest();
        xhttp.open("GET", requestUrl, true);
        xhttp.send();
        
        xhttp.onload = function() {
          searchStatus.empty();
          if (xhttp.status != 200) { // analyze HTTP status of the response
            printError(searchStatus);
          } else { // show the result
            var res = JSON.parse(this.responseText);
            if (res.status===1) {
              searchSection.show();
              buildSearchResult(searchResult, res);
            }
            else {
              printError(searchStatus, res.message);
            }
          }
        };

        xhttp.onerror = function() { // only triggers if the request couldn't be made at all
          printError(searchStatus);
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
              .attr('type', 'checkbox')
              .attr('name', 'doc')
              .attr('id', "doc"+docs[i].unid)
              .val(docs[i].unid)
              .appendTo(li);
              li.append('<label class="ml-1" for="doc'+docs[i].unid+'">'+docs[i].title+'</label>');
            }
          }
          else {
            $("<p>").addClass("color-green").text("No documents found").appendTo(parent);
          }
        }
      });
    });

    // search-result submission
    jQuery(document).ready(function(){
      // When user press on search button
      $("form[name='search-result'").on('submit', function(e){
        e.preventDefault();

        // quit if none selected
        var docChecked = $("input[type=checkbox][name='doc']:checked");
        if (docChecked.length === 0) return;

        let approvalControl = document.getElementById("approval");
        let approval = approvalControl.checked;

        // get all selected values
        var docs = [];
        docChecked.each(function(){
          docs.push($(this).val());
        });

        var outputEl = $(".js-search-result");
        outputEl.html("<p class='color-blue'>... sending data (please wait) ...</p>");

        Office.context.mailbox.getCallbackTokenAsync(function(result) {
          if (result.status !== "succeeded") {
            printError(outputEl, "Error happened (accesss token was not issued), try again or contact it@metz.dk");
            return;
          }
          
          const ewsItemId = Office.context.mailbox.item.itemId;
          const itemId = Office.context.mailbox.convertToRestId(ewsItemId, Office.MailboxEnums.RestVersion.v2_0);
          const isFromSharedFolder = Office.context.mailbox.initialData.isFromSharedFolder;
          const emailAddress = Office.context.mailbox.userProfile.emailAddress;

          // shared folder
          if (isFromSharedFolder) {
            Office.context.mailbox.item.getSharedPropertiesAsync(function(result) {
              const user = result.value.targetMailbox; 
              linkMemo(itemId, docs, user, emailAddress);
            });
          }
          // private email
          else {
            const user = "me"; 
            linkMemo(itemId, docs, user, emailAddress);
          }

          function linkMemo(itemId, docs, user, emailAddress) {
            const json = {
              "itemid": itemId,
              "docs": docs,
              "user": user,
              "emailAddress": emailAddress,
              "approval": approval
            };
      
            var app = $("#app-journalize #action option:selected").val();
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
              parent.empty();
              parent.append('<p class="color-green">Mail journalized succesfully</p>');

              $.each(data.docs, function(index, value) {
                parent.append('<p><a href="'+value.url+'" target="_blank">'+value.title+'</a></p>');
              });
            }
          }
        });
      });
    });

    function printError(el, message) {
      el.empty();
      message = message || "Error happened, try again or contact it@metz.dk";
      $("<p>")
      .addClass("color-red")
      .text(message).appendTo(el);
    }

  };

})();