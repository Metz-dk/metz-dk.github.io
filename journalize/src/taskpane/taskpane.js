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
        let val = $("input[type=radio][name='doc']:checked").val();
        var requestUrl = 'https://api-dev.metz.dk/journalize/v1/link?val='+val;
        var searchEl = $(".search-result").empty();
        $.post(requestUrl, function(data) {
          confirmLink(searchEl, data);
        })
        .fail(function() {
          $("<p>")
          .addClass("color-red")
          .text("error happened, try again or contact it@metz.dk").appendTo(searchEl);
        })
      });
    });
  };

  function confirmLink(parent, data) {
    $("<p>").addClass("color-green").text(data.message).appendTo(parent);
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

})();