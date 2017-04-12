(function ($) {
    var isIE = false;

    // Initialize common propertites 
    Office.initialize = function (reason) {
        $(document).ready(function () {
            initializeMessage();
            initializeContent();

            $("#subimtButton").click(function (event) {
                insertContentControl();
            });

            $("#subimtButton_m").click(function (event) {
                updateContentControl();
            });

            $("#privacy").click(function (event) {
                window.open('https://fachenaddin.azurewebsites.net/Appstore/ContentControlHelperPrivacy.html', '', 'resizable=1,scrollbars=1,width=1500,height=1000');
            });

            $("#help").click(function (event) {
                window.open('https://fachenaddin.azurewebsites.net/Appstore/ContentControlHelperIntro.html', '', 'resizable=1,scrollbars=1,width=1500,height=1000');
            });

            $(".clearable").on('input', function (event) {
                showClearBtn(event);
            });

            // Click clear event
            $(".cc_trash,.cc_content_trash").click(function () {
                var t = this;
                var targetId = new String(this.id);
                var textId = targetId.substring(0, targetId.length - 5);
                document.getElementById(textId).value = "";
                $("#" + targetId).hide();
            });

            $("#cc_contenttype").change(function () {
                var curValue = $(this).children('option:selected').val();
                switch (curValue) {
                    case "Text":
                        $("#PlainTextContent").show();
                        $("#HTMLContent").hide();
                        $("#OOXMLContent").hide();
                        break;

                    case "HTML":
                        $("#PlainTextContent").hide();
                        $("#HTMLContent").show();
                        $("#OOXMLContent").hide();
                        break;

                    case "OOXML":
                        $("#PlainTextContent").hide();
                        $("#HTMLContent").hide();
                        $("#OOXMLContent").show();
                        getOoxml();
                        break;

                    case "NotSet":
                        $("#PlainTextContent").hide();
                        $("#HTMLContent").hide();
                        $("#OOXMLContent").hide();
                        console.log("Insert content value not set.");
                        break;

                    default:
                        console.log("Sorry, we are out of " + curValue + ".");
                }
            });

            // Initialize IE brower showup picker
            browerShowPicker();

            Office.context.document.addHandlerAsync(
              Office.EventType.DocumentSelectionChanged, ensureSelectedRange, function (asyncResult) {
                  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                      write("addHandlerAsync in initialize: " + asyncResult.error.message);
                  }
              });
        });
    };

    // Append message to body
    function initializeMessage() {
        $('body').append(
                    '<div id="notification-message">' +
                        '<div class="padding">' +
                            '<div id="notification-message-close"></div>' +
                            '<div id="notification-message-header"></div>' +
                            '<div id="notification-message-body"></div>' +
                        '</div>' +
                    '</div>');
        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });
    }

    function showMessage(header, text) {
        // After initialization, expose a common notification function
        $('#notification-message-header').text(header);
        $('#notification-message-body').text(text);
        $('#notification-message').slideDown('fast');
    }

    // Internet Explore version check
    function isIEBrowser() {
        var Sys = {};
        var ua = navigator.userAgent.toLowerCase();
        var s;
        (s = ua.match(/msie ([\d.]+)/)) ? Sys.ie = s[1] :
        (s = ua.match(/firefox\/([\d.]+)/)) ? Sys.firefox = s[1] :
        (s = ua.match(/chrome\/([\d.]+)/)) ? Sys.chrome = s[1] :
        (s = ua.match(/opera.([\d.]+)/)) ? Sys.opera = s[1] :
        (s = ua.match(/version\/([\d.]+).*safari/)) ? Sys.safari = s[1] : 0;
        if (Sys.chrome || Sys.firefox || Sys.opera || Sys.safari) {
            isIE = false;
            return false;
        } else {
            isIE = true;
            return true;
        }
    }

    // Initialize IE version show up
    function browerShowPicker() {
        if (isIEBrowser()) {
            $(".isIEBrower").show();
            $(".notIEBrower").hide();
        }
    }

    // Show clear btn
    function showClearBtn(event) {
        var targetId = event.target.id + "trash";
        var doc = document.getElementById(targetId);
        $("#" + targetId).show();
    }

    // Extract getOoxml alone case it needs more time, make it more lazy load
    function getOoxml() {
        // Set loading status
        document.getElementById("cc_ooxml").value = "loading...";
        $("#cc_ooxml").addClass("font_change");

        // Run a batch operation
        Word.run(function (context) {
            var range = context.document.getSelection();

            var parentControl = range.parentContentControl;
            context.load(parentControl);

            return context.sync().then(function () {
                var ooxmlValue = "";
                if (!parentControl.isNull) {
                    // Set ooxml
                    ooxmlValue = parentControl.getOoxml();
                    return context.sync().then(function () {
                        document.getElementById("cc_ooxml").value = ooxmlValue.value;
                        // Reset font color
                        $("#cc_ooxml").removeClass("font_change")
                    });
                } else {
                    ooxmlValue = range.getOoxml();
                    return context.sync().then(function () {
                        document.getElementById("cc_ooxml").value = ooxmlValue.value;
                        // Reset font color
                        $("#cc_ooxml").removeClass("font_change");
                    });
                }
            });
        })
        .catch(function (error) {

            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }

    // Initialize content
    function initializeContent() {
        // Run a batch operation
        Word.run(function (context) {
            var range = context.document.getSelection();
            var parentControl = range.parentContentControl;
            context.load(parentControl, 'cannotDelete,cannotEdit,color,removeWhenEdited,style,tag,text,title,id,appearance,placeholderText,font');

            return context.sync().then(function () {
                prepareContent(parentControl, context);
                setFontPropertites(parentControl, context);
            });
        })
         .catch(function (error) {
             // Initialize controls
             initializePane();
             console.log('Error: ' + JSON.stringify(error));
             if (error instanceof OfficeExtension.Error) {
                 console.log('Debug info: ' + JSON.stringify(error.debugInfo));
             }
         });
    }

    // Catch Selection Changed
    function ensureSelectedRange() {
        //Run a batch operation
        Word.run(function (context) {
            document.getElementById("cc_ooxml").value = "";
            document.getElementById("cc_plaintext").value = "";
            document.getElementById("cc_html").value = "";
            document.getElementById("cc_contenttype").selectedIndex = 0;
            $("#PlainTextContent").hide();
            $("#HTMLContent").hide();
            $("#OOXMLContent").hide();

            var range = context.document.getSelection();
            var parentControl = range.parentContentControl;
            context.load(parentControl, 'cannotDelete,cannotEdit,color,removeWhenEdited,style,tag,text,title,id,appearance,placeholderText');

            return context.sync().then(function () {
                prepareContent(parentControl, context);
            });
        })
        .catch(function (error) {
            initializePane();
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }

    // Update ContentControl
    function updateContentControl() {
        Word.run(function (context) {
            var range = context.document.getSelection();
            context.load(range);
            return context.sync().then(function () {
                var parentControl = range.parentContentControl;
                context.load(parentControl);
                return context.sync().then(function () {
                        parentControl.tag = document.getElementById("cc_tag").value;
                        parentControl.title = document.getElementById("cc_title").value;
                        // Set locking propertites
                        var editableValue = $("#cc_editable").children('option:selected').val();
                        switch (editableValue) {
                            case "True": parentControl.cannotEdit = true; break;
                            case "False": parentControl.cannotEdit = false; break;
                            case "NotSet": console.log("Editable not set."); break;
                            default: console.log("Sorry, we are out of " + editableValue + ".")
                        }

                        var deletableValue = $("#cc_deletable").children('option:selected').val();
                        switch (deletableValue) {
                            case "True": parentControl.cannotDelete = true; break;
                            case "False": parentControl.cannotDelete = false; break;
                            case "NotSet": console.log("Deletable not set."); break;
                            default: console.log("Sorry, we are out of " + deletableValue + ".")
                        }

                        var removableValue = $("#cc_removable").children('option:selected').val();
                        switch (removableValue) {
                            case "True": parentControl.removeWhenEdited = true; break;
                            case "False": parentControl.removeWhenEdited = false; break;
                            case "NotSet": console.log("Deletable not set."); break;
                            default: console.log("Sorry, we are out of " + removableValue + ".")
                        }

                        // Set apperance related propertites
                        var colorValue = "NotSet";
                        if (!isIE) {
                            colorValue = document.getElementById('colorPan').value;
                        } else {
                            colorValue = document.getElementById('cc_colorPicker').value;
                            console.log("update ie color" + colorValue);
                        }

                        // Set font related propertites
                        var fontColorValue = "NotSet";
                        if (!isIE) {
                            fontColorValue = document.getElementById('fontcolorPan').value;
                        } else {
                            fontColorValue = document.getElementById('cc_fontcolorPicker').value;
                            console.log("update ie font color" + colorValue);
                        }

                        var fontNameValue = document.getElementById("cc_fontname").value;
                        var fontSizeValue = document.getElementById("cc_fontsize").value;

                        var styleValue = document.getElementById("cc_style").value;
                        var appearanceValue = document.getElementById("cc_showAs").value;
                        var contentType = $("#cc_contenttype").children('option:selected').val();
                        var contentValue = "";
                        switch (contentType) {
                            case "Text":
                                contentValue = document.getElementById("cc_plaintext").value;
                                parentControl.insertText(contentValue, 'replace');
                                break;

                            case "HTML":
                                contentValue = document.getElementById("cc_html").value;
                                parentControl.insertHtml(contentValue, 'replace');
                                break;

                            case "OOXML":
                                contentValue = document.getElementById("cc_ooxml").value;
                                parentControl.insertOoxml(contentValue, 'replace');
                                break;

                            case "NotSet":
                                console.log("Insert content value not set.");
                                break;

                            default:
                                console.log("Sorry, we are out of " + contentType + ".");
                        }

                        if (appearanceValue != "NotSet") {
                            parentControl.appearance = appearanceValue;
                        }

                        if (colorValue != "NotSet") {
                            parentControl.color = colorValue == "#000000" ? "#1F1F1F" : colorValue;
                        }

                        if (styleValue != "NotSet") {
                            parentControl.style = styleValue;
                        }

                        if (fontColorValue != "NotSet") {
                            parentControl.font.color = fontColorValue == "#000000" ? "#1F1F1F" : fontColorValue;
                        }

                        if (fontNameValue != "NotSet") {
                            parentControl.font.name = fontNameValue;
                        }

                        if (fontSizeValue != "NotSet") {
                            parentControl.font.size = parseFloat(fontSizeValue);
                        }

                        return context.sync().then(function () {
                            console.log('sync content control property..');
                            console.log("set font propertites done!")

                            parentControl.select();
                            showMessage("Update content control", "Success!");
                        });
                });
            });
        })
      .catch(function (error) {
          showMessage("Update content control", "Failed!");
          console.log('Error: ' + JSON.stringify(error));
          if (error instanceof OfficeExtension.Error) {
              console.log('Debug info: ' + JSON.stringify(error.debugInfo));
          }
      });
    }

    // Reads data from current document and insert a content conrol
    function insertContentControl() {

        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            var range = context.document.getSelection();
            var myContentControl = range.insertContentControl();
            var tagValue = document.getElementById("cc_tag").value;
            if (tagValue == "") {
                tagValue = document.getElementById("cc_tag").placeholder;
            }
            var titleVaule = document.getElementById("cc_title").value;
            if (titleVaule == "") {
                titleVaule = document.getElementById("cc_title").placeholder;
            }
            var placeHolderValue = document.getElementById("cc_placeholder").value;
            if (placeHolderValue == "") {
                placeHolderValue = document.getElementById("cc_placeholder").placeholder;
            }

            var editableValue = $("#cc_editable").children('option:selected').val();
            switch (editableValue) {
                case "True": myContentControl.cannotEdit = true; break;
                case "False": myContentControl.cannotEdit = false; break;
                case "NotSet": console.log("Editable not set."); break;
                default: console.log("Sorry, we are out of " + editableValue + ".")
            }

            var deletableValue = $("#cc_deletable").children('option:selected').val();
            switch (deletableValue) {
                case "True": myContentControl.cannotDelete = true; break;
                case "False": myContentControl.cannotDelete = false; break;
                case "NotSet": console.log("Deletable not set."); break;
                default: console.log("Sorry, we are out of " + deletableValue + ".")
            }

            var removableValue = $("#cc_removable").children('option:selected').val();
            switch (removableValue) {
                case "True": myContentControl.removeWhenEdited = true; break;
                case "False": myContentControl.removeWhenEdited = false; break;
                case "NotSet": console.log("Deletable not set."); break;
                default: console.log("Sorry, we are out of " + removableValue + ".")
            }

            // Set outlook propertites
            var colorValue = "NotSet";
            var fontColorValue = "NotSet";
            if (!isIE) {
                colorValue = document.getElementById('colorPan').value;
                fontColorValue = document.getElementById('fontcolorPan').value;
            } else {
                colorValue = document.getElementById('cc_colorPicker').value;
                fontColorValue = document.getElementById('cc_fontcolorPicker').value;
                console.log("insert ie color" + colorValue);
                console.log("insert ie font color" + fontColorValue);
            }

            var styleValue = document.getElementById("cc_style").value;
            var appearanceValue = document.getElementById("cc_showAs").value;
            var fontSizeValue = document.getElementById("cc_fontsize").value;
            var fontNameValue = document.getElementById('cc_fontname').value;

            // Insert selected content type value
            var contentType = $("#cc_contenttype").children('option:selected').val();
            var contentValue = "";
            switch (contentType) {
                case "Text":
                    contentValue = document.getElementById("cc_plaintext").value;
                    myContentControl.insertText(contentValue, 'replace');
                    break;

                case "HTML":
                    contentValue = document.getElementById("cc_html").value;
                    myContentControl.insertHtml(contentValue, 'replace');
                    break;

                case "OOXML":
                    contentValue = document.getElementById("cc_ooxml").value;
                    myContentControl.insertOoxml(contentValue, 'replace');
                    break;

                case "NotSet":
                    console.log("Insert content value not set.");
                    break;

                default:
                    console.log("Sorry, we are out of " + contentType + ".");
            }

            // Set common propertites
            myContentControl.tag = tagValue;
            myContentControl.title = titleVaule;

            if (appearanceValue != "NotSet") {
                myContentControl.appearance = appearanceValue;
            }

            if (colorValue != "NotSet") {
                myContentControl.color = colorValue == "#000000" ? "#1F1F1F" : colorValue;
            }

            if (fontColorValue != "NotSet") {
                myContentControl.font.color = fontColorValue == "#000000" ? "#1F1F1F" : fontColorValue;
            }

            if (fontSizeValue != "NotSet") {
                myContentControl.font.size = parseFloat(fontSizeValue);
            }

            if (fontNameValue != "NotSet") {
                myContentControl.font.name = fontNameValue;
            }

            if (styleValue != "NotSet") {
                myContentControl.style = styleValue;
            }

            myContentControl.select();

            // Load myContentControl for display 'Update' UI
            context.load(myContentControl);
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log("Start reInitialize content control, show 'Update' UI");
                setTimeout(function () {
                    return context.sync().then(function () {
                        document.getElementById("cc_ooxml").value = "";
                        document.getElementById("cc_plaintext").value = "";
                        document.getElementById("cc_html").value = "";
                        document.getElementById("cc_contenttype").selectedIndex = 0;
                        $("#PlainTextContent").hide();
                        $("#HTMLContent").hide();
                        $("#OOXMLContent").hide();

                        if (!myContentControl.isNull) {
                            $("#div_o").hide();
                            $("#div_m").show();
                            $("#readonlyId").show();
                            $(".font_property").addClass("ms-Toggle-description");
                            if (myContentControl.tag == null) {
                                document.getElementById("cc_tag").value = "";
                            } else {
                                document.getElementById("cc_tag").value = myContentControl.tag;
                            }
                            if (myContentControl.title == null) {
                                document.getElementById("cc_title").value = "";
                            } else {
                                document.getElementById("cc_title").value = myContentControl.title;
                            }

                            document.getElementById("cc_id").value = myContentControl.id;
                            document.getElementById("cc_plaintext").value = myContentControl.text;

                            // Set locking propertites
                            var editableValue = myContentControl.cannotEdit;
                            if (editableValue) {
                                document.getElementById("cc_editable").selectedIndex = 1;
                            } else {
                                document.getElementById("cc_editable").selectedIndex = 2;
                            }

                            var deletableValue = myContentControl.cannotDelete;
                            if (deletableValue) {
                                document.getElementById("cc_deletable").selectedIndex = 1;
                            } else {
                                document.getElementById("cc_deletable").selectedIndex = 2;
                            }

                            var removableValue = myContentControl.removeWhenEdited;
                            if (removableValue) {
                                document.getElementById("cc_removable").selectedIndex = 1;
                            } else {
                                document.getElementById("cc_removable").selectedIndex = 2;
                            }

                            // Set selected options
                            // Set color
                            var color = myContentControl.color;
                            if (!isIE) {
                                document.getElementById('colorPan').value = color;
                            } else {
                                document.getElementById('cc_colorPicker').value = color;
                                $("#cc_colorPicker").css("background-color", color);
                                console.log("initialize ie color " + color);
                            }

                            // Set style
                            var isFound = false;
                            var style = myContentControl.style;
                            var styleDoc = document.getElementById("cc_style");
                            for (sidx = 0; sidx < styleDoc.length; sidx++) {
                                if (styleDoc[sidx].textContent == style) {
                                    styleDoc.selectedIndex = sidx;
                                    isFound = true;
                                    break;
                                }
                            }
                            if (!isFound) {
                                $("#cc_style").append("<option>" + style + "</option>");
                                styleDoc.selectedIndex = styleDoc.length - 1;
                            }

                            isFound = false;
                            var appearance = myContentControl.appearance;
                            var appearanceDoc = document.getElementById("cc_showAs");
                            for (sidx = 0; sidx < appearanceDoc.length; sidx++) {
                                if (appearanceDoc[sidx].textContent == appearance) {
                                    appearanceDoc.selectedIndex = sidx;
                                    isFound = true;
                                    break;
                                }
                            }
                            if (!isFound) {
                                $("#cc_showAs").append("<option>" + appearance + "</option>");
                                appearanceDoc.selectedIndex = appearanceDoc.length - 1;
                            }

                            var fontColor = fontColorValue == "#000000" ? "#1F1F1F" : fontColorValue;
                            if (!isIE) {
                                document.getElementById('fontcolorPan').value = fontColor;
                            } else {
                                fontColor = "#" + fontColor;
                                document.getElementById('cc_fontcolorPicker').value = fontColor;
                                $("#cc_fontcolorPicker").css("background-color", fontColor);
                                console.log("initialize ie font color " + fontColor);
                            }

                            isFound = false;
                            var fontSize = fontSizeValue;
                            var fontSizeDoc = document.getElementById("cc_fontsize");
                            for (sidx = 0; sidx < fontSizeDoc.length; sidx++) {
                                if (fontSizeDoc[sidx].textContent == fontSize) {
                                    fontSizeDoc.selectedIndex = sidx;
                                    isFound = true;
                                    break;
                                }
                            }
                            if (!isFound) {
                                $("#cc_fontsize").append("<option>" + fontSize + "</option>");
                                fontSizeDoc.selectedIndex = fontSizeDoc.length - 1;
                            }

                            isFound = false;
                            var fontName = fontNameValue;
                            var fontNameDoc = document.getElementById("cc_fontname");
                            for (sidx = 0; sidx < fontNameDoc.length; sidx++) {
                                if (fontNameDoc[sidx].textContent == fontName) {
                                    fontNameDoc.selectedIndex = sidx;
                                    isFound = true;
                                    break;
                                }
                            }
                            if (!isFound) {
                                $("#cc_fontname").append("<option>" + fontName + "</option>");
                                fontNameDoc.selectedIndex = fontNameDoc.length - 1;
                            }
                            showMessage("Insert content control", "Success!");
                            console.log('Insert Wrapped a content control around the selected text.');
                        } else {
                            showMessage("Insert content control", "Failed!");
                            console.log("Error, insert content control failed.");
                        }
                    });
                }, 300);
            });
        })
        .catch(function (error) {
            showMessage("Insert content control", "Failed!");
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }

    function prepareContent(parentControl, context) {
        // Initialize controls
        document.getElementById("cc_ooxml").value = "";
        document.getElementById("cc_plaintext").value = "";
        document.getElementById("cc_html").value = "";
        document.getElementById("cc_contenttype").selectedIndex = 0;
        $("#PlainTextContent").hide();
        $("#HTMLContent").hide();
        $("#OOXMLContent").hide();

            $("#div_o").hide();
            $("#div_m").show();
            $("#readonlyId").show();
            $(".font_property").addClass("ms-Toggle-description");
            if (parentControl.tag == null) {
                document.getElementById("cc_tag").value = "";
            } else {
                document.getElementById("cc_tag").value = parentControl.tag;
            }
            if (parentControl.title == null) {
                document.getElementById("cc_title").value = "";
            } else {
                document.getElementById("cc_title").value = parentControl.title;
            }

            document.getElementById("cc_placeholder").value = parentControl.placeholderText;
            document.getElementById("cc_id").value = parentControl.id;
            document.getElementById("cc_plaintext").value = parentControl.text;

            // Set locking propertites
            var editableValue = parentControl.cannotEdit;
            if (editableValue) {
                document.getElementById("cc_editable").selectedIndex = 1;
            } else {
                document.getElementById("cc_editable").selectedIndex = 2;
            }

            var deletableValue = parentControl.cannotDelete;
            if (deletableValue) {
                document.getElementById("cc_deletable").selectedIndex = 1;
            } else {
                document.getElementById("cc_deletable").selectedIndex = 2;
            }

            var removableValue = parentControl.removeWhenEdited;
            if (removableValue) {
                document.getElementById("cc_removable").selectedIndex = 1;
            } else {
                document.getElementById("cc_removable").selectedIndex = 2;
            }

            // Set selected options
            // Set color
            var color = parentControl.color;
            if (!isIE) {
                document.getElementById('colorPan').value = color;
            } else {
                document.getElementById('cc_colorPicker').value = color;
                $("#cc_colorPicker").css("background-color", color);
                console.log("initialize ie color " + color);
            }

            // Set style
            var isFound = false;
            var style = parentControl.style;
            var styleDoc = document.getElementById("cc_style");
            for (sidx = 0; sidx < styleDoc.length; sidx++) {
                if (styleDoc[sidx].textContent == style) {
                    styleDoc.selectedIndex = sidx;
                    isFound = true;
                    break;
                }
            }

            if (!isFound) {
                $("#cc_style").append("<option>" + style + "</option>");
                styleDoc.selectedIndex = styleDoc.length - 1;
            }

            // Set Appearance
            isFound = false;
            var appearance = parentControl.appearance;
            var appearanceDoc = document.getElementById("cc_showAs");
            for (sidx = 0; sidx < appearanceDoc.length; sidx++) {
                if (appearanceDoc[sidx].textContent == appearance) {
                    appearanceDoc.selectedIndex = sidx;
                    isFound = true;
                    break;
                }
            }

            if (!isFound) {
                $("#cc_showAs").append("<option>" + appearance + "</option>");
                appearanceDoc.selectedIndex = appearanceDoc.length - 1;
            }
    }

    function setFontPropertites(parentControl, context) {
        var font = parentControl.font;
        context.load(font);

        return context.sync().then(function () {
            // Set font color
            var fontColor = font.color;
            if (!isIE) {
                document.getElementById('fontcolorPan').value = fontColor;
            } else {
                document.getElementById('cc_fontcolorPicker').value = fontColor;
                $("#cc_fontcolorPicker").css("background-color", fontColor);
                console.log("initialize ie font color " + fontColor);
            }

            // Set font size
            isFound = false;
            var fontSize = font.size;
            var fontSizeDoc = document.getElementById("cc_fontsize");
            for (sidx = 0; sidx < fontSizeDoc.length; sidx++) {
                if (fontSizeDoc[sidx].textContent == fontSize) {
                    fontSizeDoc.selectedIndex = sidx;
                    isFound = true;
                    break;
                }
            }

            if (!isFound) {
                $("#cc_fontsize").append("<option>" + fontSize + "</option>");
                fontSizeDoc.selectedIndex = fontSizeDoc.length - 1;
            }

            // Set font name
            isFound = false;
            var fontName = font.name;
            var fontNameDoc = document.getElementById("cc_fontname");
            for (sidx = 0; sidx < fontNameDoc.length; sidx++) {
                if (fontNameDoc[sidx].textContent == fontName) {
                    fontNameDoc.selectedIndex = sidx;
                    isFound = true;
                    break;
                }
            }

            if (!isFound) {
                $("#cc_fontname").append("<option>" + fontName + "</option>");
                fontNameDoc.selectedIndex = fontNameDoc.length - 1;
            }
        });
    }

    function initializePane() {
        $("#div_o").show();
        $("#div_m").hide();
        $("#readonlyId").hide();
        $(".cc_trash,.cc_content_trash").hide();
        $(".font_property").removeClass("ms-Toggle-description");
        $("#cc_colorPicker").css("background-color", "#1F1F1F");
        $("#cc_fontcolorPicker").css("background-color", "#1F1F1F");
        document.getElementById("cc_tag").value = "";
        document.getElementById("cc_title").value = "";
        document.getElementById("cc_placeholder").value = "";
        document.getElementById('colorPan').value = "#1F1F1F";
        document.getElementById('fontcolorPan').value = "#1F1F1F";
        document.getElementById('cc_colorPicker').value = "1F1F1F";
        document.getElementById('cc_fontcolorPicker').value = "1F1F1F";
        document.getElementById("cc_fontname").selectedIndex = 0;
        document.getElementById("cc_fontsize").selectedIndex = 0;
        document.getElementById("cc_showAs").selectedIndex = 0;
        document.getElementById("cc_style").selectedIndex = 0;
        document.getElementById("cc_contenttype").selectedIndex = 0;
        document.getElementById("cc_editable").selectedIndex = 0;
        document.getElementById("cc_removable").selectedIndex = 0;
        document.getElementById("cc_deletable").selectedIndex = 0;
    }

})(jQuery);