/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

(function(){
  'use strict';

  // Cache for the full result object
  let cachedResult = null;

  // Optimized token retrieval function with caching
  function getCallbackTokenWithRetry(callback) {
    // If we have a cached result, use it regardless of status
    if (cachedResult) {
      callback(cachedResult);
      return;
    }
    
    Office.context.mailbox.getCallbackTokenAsync(function(result) {
      // Cache the entire result object
      cachedResult = result;
      callback(result);
    });
  }

  // Improved sendRequest function using fetch with timeout
  function sendRequest(endpoint, data, options = {}) {
    const { method = 'POST', timeout = 30000 } = options;
    
    // Create abort controller for timeout
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeout);
    
    const fetchOptions = {
      method,
      headers: {
        'Content-Type': 'application/json'
      },
      signal: controller.signal
    };
    
    // Add body for POST requests
    if (method === 'POST' && data) {
      fetchOptions.body = JSON.stringify(data);
    }
    
    // For GET requests with data, append to URL
    let url = endpoint;
    if (method === 'GET' && data) {
      const params = new URLSearchParams();
      Object.entries(data).forEach(([key, value]) => {
        params.append(key, value);
      });
      url = `${endpoint}?${params.toString()}`;
    }
    
    return fetch(url, fetchOptions)
      .then(response => {
        clearTimeout(timeoutId);
        if (!response.ok) {
          throw new Error(`HTTP error! Status: ${response.status}`);
        }
        return response.json();
      })
      .catch(error => {
        clearTimeout(timeoutId);
        // If it's a timeout, provide a clearer error
        if (error.name === 'AbortError') {
          throw new Error('Request timed out');
        }
        throw error;
      });
  }

  Office.initialize = function(reason){
    document.addEventListener('DOMContentLoaded', function() {
      const validationStatus = document.querySelector(".js-search-status");
      if (validationStatus) validationStatus.textContent = '';
    
      getCallbackTokenWithRetry(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          const errorMessage = result.error ? result.error.message : "Unknown error";
          console.log("Token error details:", result.error);
          printError(validationStatus, errorMessage + " - Error (9001) (Try again or contact it@metz.dk)");
          return;
        }

        const ewsItemId = Office.context.mailbox.item.itemId;
        const itemId = Office.context.mailbox.convertToRestId(ewsItemId, Office.MailboxEnums.RestVersion.v2_0);
        const isFromSharedFolder = Office.context.mailbox.initialData.isFromSharedFolder;
        const emailAddress = Office.context.mailbox.userProfile.emailAddress;
    
        if (isFromSharedFolder) {
          Office.context.mailbox.item.getSharedPropertiesAsync(function(result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              printError(validationStatus, "Failed to get shared properties - Error (9003) (Try again or contact it@metz.dk)");
              return;
            }
            validateMemo(itemId, result.value.targetMailbox, emailAddress, validationStatus);
          });
        } else {
          validateMemo(itemId, "me", emailAddress, validationStatus);
        }
      });
    
      function validateMemo(itemId, user, emailAddress, validationStatus) {
        const endpoint = "https://api.metz.dk/journalize/v1/validate";
        const data = {
          itemid: itemId,
          user: user,
          emailAddress: emailAddress,
        };
    
        sendRequest(endpoint, data)
          .then(res => {
            if (validationStatus) validationStatus.textContent = '';
            togglePaneControls(!res.valid);
            if (!res.valid) {
              printError(validationStatus, (res.message || "Validation error"));
            }
          })
          .catch(error => {
            console.error("Validation request failed:", error);
            togglePaneControls(false);
            printError(validationStatus, "Request failed: " + error.message);
          });
      }
    
      function togglePaneControls(flag) {
        const formElements = document.querySelectorAll("form[name='search'] input, form[name='search'] select, form[name='search'] button");
        formElements.forEach(element => {
          element.disabled = flag;
        });
      }
    });

    // search
    document.addEventListener('DOMContentLoaded', function() {
      const selectElement = document.querySelector("form[name='search'] select");
      if (selectElement) selectElement.focus();

      // When user changes the action selection
      document.querySelector("form[name='search'] select")?.addEventListener('change', function() {
        const action = document.getElementById('action')?.value;
        const approvalContainer = document.querySelector('.js-approval');
        
        if (!approvalContainer) return;
        
        const flag = action === 'order-open' || action === 'order-closed' || action === 'order-rma';
        approvalContainer.style.display = flag ? 'block' : 'none';

        // Reset the approval checkbox
        const approvalCheckbox = document.getElementById('approval');
        if (approvalCheckbox) approvalCheckbox.checked = false;
      });

      // When user press on search button
      document.querySelector("form[name='search']")?.addEventListener('submit', function(e) {
        e.preventDefault();

        const action = document.querySelector('#action')?.value;
        const keyword = document.getElementById("SearchQuery")?.value;
        const recentControl = document.getElementById("recent");
        const recent = recentControl?.checked ? "1" : "0";

        if (!action || !keyword) return;

        const searchData = { action, keyword, recent };
        const searchSection = document.querySelector(".js-search-section");
        const searchResult = document.querySelector(".js-search-result");
        const searchStatus = document.querySelector(".js-search-status");
        
        if (searchSection) searchSection.style.display = 'none';
        if (searchResult) searchResult.textContent = '';
        if (searchStatus) {
          searchStatus.textContent = '';
          const waitMessage = document.createElement('p');
          waitMessage.className = 'color-blue';
          waitMessage.textContent = '... please wait...';
          searchStatus.appendChild(waitMessage);
        }

        sendRequest('https://api.metz.dk/journalize/v1/search', searchData, { method: 'GET' })
          .then(res => {
            if (searchStatus) searchStatus.textContent = '';
            
            if (res.status === 1) {
              if (searchSection) searchSection.style.display = 'block';
              buildSearchResult(searchResult, res);
            } else {
              printError(searchStatus, res.message);
            }
          })
          .catch(error => {
            console.error("Search request failed:", error);
            printError(searchStatus, "Search failed: " + error.message);
          });

        function buildSearchResult(parent, data) {
          if (!parent) return;
          
          parent.textContent = '';
          const action = data.action;
          const docs = data.docs;
      
          if (docs.length > 0) {
            const statusP = document.createElement('p');
            statusP.className = 'color-green';
            statusP.textContent = docs.length + " document(s) displayed (total: " + data.total + ")";
            parent.appendChild(statusP);
            
            const list = document.createElement('ul');
            list.className = 'my-3';
            parent.appendChild(list);
            
            for (let i = 0; i < docs.length; i++) {
              const li = document.createElement('li');
              list.appendChild(li);
              
              const input = document.createElement('input');
              input.type = 'checkbox';
              input.name = 'doc';
              input.id = "doc" + docs[i].unid;
              input.value = docs[i].unid;
              li.appendChild(input);
              
              const label = document.createElement('label');
              label.className = 'ml-1';
              label.htmlFor = "doc" + docs[i].unid;
              label.textContent = docs[i].title;
              li.appendChild(label);
            }
          } else {
            const noDocsP = document.createElement('p');
            noDocsP.className = 'color-green';
            noDocsP.textContent = "No documents found";
            parent.appendChild(noDocsP);
          }
        }
      });
    });

    // search-result submission
    document.addEventListener('DOMContentLoaded', function() {
      // When user press on search button
      document.querySelector("form[name='search-result']")?.addEventListener('submit', function(e) {
        e.preventDefault();

        // quit if none selected
        const docChecked = document.querySelectorAll("input[type=checkbox][name='doc']:checked");
        if (docChecked.length === 0) return;

        const approvalControl = document.getElementById("approval");
        const approval = approvalControl?.checked;

        // get all selected values
        const docs = Array.from(docChecked).map(checkbox => checkbox.value);

        const outputEl = document.querySelector(".js-search-result");
        if (outputEl) {
          outputEl.textContent = '';
          const waitMessage = document.createElement('p');
          waitMessage.className = 'color-blue';
          waitMessage.textContent = '... sending data (please wait) ...';
          outputEl.appendChild(waitMessage);
        }

        getCallbackTokenWithRetry(function(result) {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            const errorMessage = result.error ? result.error.message : "Unknown error";
            console.log("Token error details:", result.error);
            printError(outputEl, errorMessage + " - Error (9002) (Try again or contact it@metz.dk)");
            // If we get an error, clear the cache so next attempt gets a fresh token
            cachedResult = null;
            return;
          }
          
          const ewsItemId = Office.context.mailbox.item.itemId;
          const itemId = Office.context.mailbox.convertToRestId(ewsItemId, Office.MailboxEnums.RestVersion.v2_0);
          const isFromSharedFolder = Office.context.mailbox.initialData.isFromSharedFolder;
          const emailAddress = Office.context.mailbox.userProfile.emailAddress;

          // shared folder
          if (isFromSharedFolder) {
            Office.context.mailbox.item.getSharedPropertiesAsync(function(result) {
              if (result.status !== Office.AsyncResultStatus.Succeeded) {
                printError(outputEl, "Failed to get shared properties - Error (9004) (Try again or contact it@metz.dk)");
                return;
              }
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
      
            const actionSelect = document.querySelector("#app-journalize #action");
            const app = actionSelect ? actionSelect.value : '';
            if (!app) {
              printError(outputEl, "Failed to determine action");
              return;
            }
            
            const endpoint = "https://api.metz.dk/journalize/v1/" + app;
            
            sendRequest(endpoint, json)
              .then(res => {
                if (res.status === 1) {
                  confirmLink(outputEl, res);
                } else {
                  printError(outputEl, res.message || "Unknown error");
                }
              })
              .catch(error => {
                console.error("Journalize request failed:", error);
                printError(outputEl, "Request failed: " + error.message);
              });
          
            function confirmLink(parent, data) {
              if (!parent) return;
              
              parent.textContent = '';
              
              const successP = document.createElement('p');
              successP.className = 'color-green';
              successP.textContent = 'Mail journalized successfully';
              parent.appendChild(successP);

              if (data.docs && Array.isArray(data.docs)) {
                data.docs.forEach(doc => {
                  const linkP = document.createElement('p');
                  const link = document.createElement('a');
                  link.href = doc.url;
                  link.target = '_blank';
                  link.textContent = doc.title;
                  linkP.appendChild(link);
                  parent.appendChild(linkP);
                });
              }
            }
          }
        });
      });
    });

    function printError(el, message) {
      if (!el) return;
      
      el.textContent = '';
      message = message || "Error happened, try again or contact it@metz.dk";
      
      const errorP = document.createElement('p');
      errorP.className = 'color-red';
      errorP.textContent = message;
      el.appendChild(errorP);
    }
  };
})();