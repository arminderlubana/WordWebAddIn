
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
           // var element = document.querySelector('.ms-MessageBanner');
           // messageBanner = new fabric.MessageBanner(element);
           // messageBanner.hideBanner();

           // // If not using Word 2016, use fallback logic.
           // if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
           //     $("#template-description").text("This sample displays the selected text.");
           //     $('#button-text').text("Display!");
           //     $('#button-desc').text("Display the selected text");
                
           //    // $('#highlight-button').click(displaySelectedText);
           //     return;
           // }

           // $("#template-description").text("This sample highlights the longest word in the text you have selected in the document.");
           // $('#button-text').text("Highlight!");
           // $('#button-desc').text("Highlights the longest word.");
            
           // //loadSampleData();
           // getDocumentAsCompressed();
           //// sendFile();
           // // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightLongestWord);
           
       // $("#highlight-button").click(() => tryCatch(readCustomDocumentProperties));
         //   readCustomDocumentProperties();
           // tryCatch(readCustomDocumentProperties);
        });
    };
    function hightlightLongestWord() {
        Word.run(function (context) {
            var customDocProps = context.document.properties.customProperties;
            context.load(customDocProps);
            return context.sync()
                .then(function () {
                   // console.log(customDocProps.items.length);
                    var employmentId;
                    var companyId;
                    var url;
                    for (var i = 0; i < customDocProps.items.length; i++) {
                        console.log("Property Name:" + customDocProps.items[i].key + ";Type=" + customDocProps.items[i].type + "; Property Value=" + customDocProps.items[i].value);
                        var key = customDocProps.items[i].key;
                        switch (key) {
                            case "EmploymentId":
                                employmentId = customDocProps.items[i].value;
                                break;
                            case "CompanyId":
                                companyId = customDocProps.items[i].value;
                                break;
                            case "EndPointToUpdateContract":
                                url = customDocProps.items[i].value;
                                break;
                            default:
                        }

                    }
                    getDocumentAsCompressed(employmentId, companyId, url);
                })

            //// Queue a command to get the current selection and then
            //// create a proxy range object with the results.
            //var range = context.document.getSelection();
            
            //// This variable will keep the search results for the longest word.
            //var searchResults;
            
            //// Queue a command to load the range selection result.
            //context.load(range, 'text');

            //// Synchronize the document state by executing the queued commands
            //// and return a promise to indicate task completion.
            //return context.sync()
            //    .then(function () {
            //        // Get the longest word from the selection.
            //        var words = range.text.split(/\s+/);
            //        var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

            //        // Queue a search command.
            //        searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

            //        // Queue a commmand to load the font property of the results.
            //        context.load(searchResults, 'font');
            //    })
            //    .then(context.sync)
            //    .then(function () {
            //        // Queue a command to highlight the search results.
            //        searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
            //        searchResults.items[0].font.bold = true;
            //    })
            //    .then(context.sync);
        })
        .catch(errorHandler);
    } 
     function readCustomDocumentProperties() {
         Office.context.document.getFilePropertiesAsync(function (asyncResult) {
             var result = asyncResult;
         });
    }

    function getDocumentAsCompressed(employmentId, companyId, url) {
        Office.context.document.getFileAsync("compressed", { sliceSize: 4194304   },
            function (result) {
                if (result.status == Office.AsyncResultStatus.Succeeded) {
                    // If the getFileAsync call succeeded, then
                    // result.value will return a valid File Object.
                    var myFile = result.value;
                    //uploadFile(myFile);
                    //return;
                    var sliceCount = myFile.sliceCount;
                    var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];
                    //showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);

                    // Get the file slices.
                    getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived, employmentId, companyId, url);
                }
                else {
                    showNotification("Error:", result.error.message);
                }
            });
    }



    function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived, employmentId, companyId, url) {
        file.getSliceAsync(nextSlice, function (sliceResult) {
            if (sliceResult.status == "succeeded") {
                if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                    return;
                }

                // Got one slice, store it in a temporary array.
                // (Or you can do something else, such as
                // send it to a third-party server.)
                docdataSlices[sliceResult.value.index] = sliceResult.value.data;
                if (++slicesReceived == sliceCount) {
                    // All slices have been received.
                    file.closeAsync();
                    onGotAllSlices(docdataSlices, employmentId, companyId, url);
                }
                else {
                    getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived, employmentId, companyId, url);
                }
            }
            else {
                gotAllSlices = false;
                file.closeAsync();
                showNotification("getSliceAsync Error:", sliceResult.error.message);
            }
        });
    }

    function onGotAllSlices(docdataSlices, employmentId, companyId, url) {
        var docdata = [];
        for (var i = 0; i < docdataSlices.length; i++) {
            docdata = docdata.concat((docdataSlices[i]));
        }
       
        var fileContent = new String();
        for (var j = 0; j < docdata.length; j++) {
            fileContent += String.fromCharCode(docdata[j]);
        }

       // uploadFile(file);
       // var file = base64ToFile(docdata, "arminder.docx", "application/msword");
       // var file = converBase64toBlob(docdata, "application/msword");
        var file= getFileFromFileContent(fileContent);
       
        uploadFile(file, employmentId, companyId, url);
        // Now all the file content is stored in 'fileContent' variable,
        // you can do something with it, such as print, fax...
    }

    function getFileFromFileContent(fileContent ) {
        let bytes = new Uint8Array(fileContent.length);

        for (let i = 0; i < bytes.length; i++) {
            bytes[i] = fileContent.charCodeAt(i);
        }

        var blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        return blob;
    }
    function base64ToFile(base64Data, tempfilename, contentType) {
        contentType = contentType || '';
        var sliceSize = 1024;
        var byteCharacters = atob(base64Data);
        //var byteCharacters = base64Data;
        var bytesLength = byteCharacters.length;
        var slicesCount = Math.ceil(bytesLength / sliceSize);
        var byteArrays = new Array(slicesCount);

        for (var sliceIndex = 0; sliceIndex < slicesCount; ++sliceIndex) {
            var begin = sliceIndex * sliceSize;
            var end = Math.min(begin + sliceSize, bytesLength);

            var bytes = new Array(end - begin);
            for (var offset = begin, i = 0; offset < end; ++i, ++offset) {
                bytes[i] = byteCharacters[offset].charCodeAt(0);
            }
            byteArrays[sliceIndex] = new Uint8Array(bytes);
        }
        var file = new File(byteArrays, tempfilename, { type: contentType });
        return file;
    }

    function uploadFile(data,employmentId,companyId,url) {
        var formData = new FormData();
        //var blob = new Blob([data], { type: 'application/json' });
        //0eb4d0e9 - 460a - 4f6b - b26d - 08d61d27c5cd
        formData.append("File", data);
        var endpoint = 'https://test-everyone.westeurope.cloudapp.azure.com/api/v1/GenerateContract/UpdateContractByEmploymentId/' + employmentId + ',' + companyId;
        $.ajax({
            url: endpoint ,            
            type: 'POST',
            data: formData,           
            crossDomain: true,
            contentType: false, // NEEDED, DON'T OMIT THIS (requires jQuery 1.6+)
            processData: false,
        }).done(function (data) {
            var a = data;
            // placeholder
            }).fail(function (status) {
                var s = status;
            // placeholder
            }).always(function () {
                var c;
            // placeholder
        });    // Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
    function sendFile() {
        Office.context.document.getFileAsync("compressed",
            { sliceSize: 100000 },
            function (result) {

                if (result.status == Office.AsyncResultStatus.Succeeded) {

                    // Get the File object from the result.
                    var myFile = result.value;
                    
                    var state = {
                        file: myFile,
                        counter: 0,
                        sliceCount: myFile.sliceCount
                    };

                    updateStatus("Getting file of " + myFile.size + " bytes");
                    getSlice(state);
                }
                else {
                    updateStatus(result.status);
                }
            });
    }

    }
    function sendSlice(slice, state) {
        var data = slice.data;

        // If the slice contains data, create an HTTP request.
        if (data) {
            
            // Encode the slice data, a byte array, as a Base64 string.
            // NOTE: The implementation of myEncodeBase64(input) function isn't 
            // included with this example. For information about Base64 encoding with
            // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
          //  var fileData = btoa(data);
            var formData = new FormData();
            var blob = new Blob([fileData], { type: 'application/json' });
            formData.append("File", data, "arminder.docx");
            // Create a new HTTP request. You need to send the request 
            // to a webpage that can receive a post.
            var request = new XMLHttpRequest();

            // Create a handler function to update the status 
            // when the request has been sent.
            request.onreadystatechange = function () {
                if (request.readyState == 4) {

                    updateStatus("Sent " + slice.size + " bytes.");
                    state.counter++;

                    if (state.counter < state.sliceCount) {
                        getSlice(state);
                    }
                    else {
                        closeFile(state);
                    }
                }
            }
            request.crossDomain = true;
            //request.timeout = 5000;
            request.withCredentials = false;
            //request.open("POST", "http://dev-everyone.westeurope.cloudapp.azure.com:2200/api/v1/GenerateContract/UpdateContractByEmploymentId/53606db3-5a92-4da3-60d4-08d613e462c4,BK");
            request.open("POST", "https://test-everyone.westeurope.cloudapp.azure.com/api/v1/GenerateContract/UpdateContractByEmploymentId/53606db3-5a92-4da3-60d4-08d613e462c4,BK");
            //request.setRequestHeader("Slice-Number", slice.index);

            // Send the file as the body of an HTTP POST 
            // request to the web server.
            request.onload = function () {
                if (request.status === 200) {
                    alert('An error occurred!');
                    // File(s) uploaded.
                   // uploadButton.innerHTML = 'Upload';
                } else {
                    alert('An error occurred!');
                }
            };

            request.send(formData);
            //uploadFile(data);
          
        }
    }
   
    // Get a slice from the file and then call sendSlice.
    function getSlice(state) {
        state.file.getSliceAsync(state.counter, function (result) {
            if (result.status == Office.AsyncResultStatus.Succeeded) {
                updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
                sendSlice(result.value, state);
            }
            else {
                updateStatus(result.status);
            }
        });
    }
    function closeFile(state) {
        // Close the file when you're done with it.
        state.file.closeAsync(function (result) {

            // If the result returns as a success, the
            // file has been successfully closed.
            if (result.status == "succeeded") {
                updateStatus("File closed.");
            }
            else {
                updateStatus("File couldn't be closed.");
            }
        });
    }
    function updateStatus(message) {
        var statusInfo = $('#status');
        statusInfo.innerHTML += message + "<br/>";
    }
    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
        .catch(errorHandler);
    }

   


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    function checkForMIMEType(response) {
        var blob;
        //if (response.mimetype == 'pdf') {
        //    blob = converBase64toBlob(response.content, 'application/pdf');
        //} else if (response.mimetype == 'doc') {
            blob = converBase64toBlob(response.content, 'application/msword');
            /*Find the content types for different format of file at http://www.freeformatter.com/mime-types-list.html*/
       // }
        var blobURL = URL.createObjectURL(blob);
        window.open(blobURL);
    }
    function converBase64toBlob(content, contentType) {
        contentType = contentType || '';
        var sliceSize = 512;
        var byteCharacters = window.atob(content); //method which converts base64 to binary
        var byteArrays = [
        ];
        for (var offset = 0; offset < byteCharacters.length; offset += sliceSize) {
            var slice = byteCharacters.slice(offset, offset + sliceSize);
            var byteNumbers = new Array(slice.length);
            for (var i = 0; i < slice.length; i++) {
                byteNumbers[i] = slice.charCodeAt(i);
            }
            var byteArray = new Uint8Array(byteNumbers);
            byteArrays.push(byteArray);
        }
        var blob = new Blob(byteArrays, {
            type: contentType
        }); //statement which creates the blob
        return blob;
    }

})();
