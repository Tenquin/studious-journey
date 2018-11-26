//This JS file is custom for this application.

//if you want code to run as soon as document load, put them in the below funciton
$(document).ready(function () {
    jHideAllObjects();
    jqEvents();
    //btnRecover.value = "Recover "  + curModel;
    $('[data-toggle="tooltip"]').tooltip();
    MainLoop();
    Window_OnLoad();

});

'if you want to globally hide all objects put the in this function and call them in your main.vbs states'
function jHideAllObjects() {
    jHide_DivButtons();
    jHide_DivPowerPanel();
    //jHide_BtnCaptueWIM();
}


// put events here which you can assign to your HTML objects
function jqEvents() {
    $("#gear1").on("click", function () {
        //reset the msgState
        msgState = "INIT"
    });

 
} // end of events


function jShow_DivButtons() {
    $("#divButtons").slideDown(1000);
}

function jHide_DivButtons() {
    $("#divButtons").slideUp(1000);
}

function jShow_DivPowerPanel() {
    $("#divPowerPanel").slideDown(1000);
}

function jHide_DivPowerPanel() {
    $("#divPowerPanel").slideUp(1000);
}

function jSpecialSlide(el, properties, options) {
    el.css({
        visibility: 'hidden', // Hide it so the position change isn't visible
        position: 'absolute'
    });
    var origPos = el.position();
    el.css({
        visibility: 'visible', // Unhide it (now that position is 'absolute')
        position: 'absolute',
        top: origPos.top,
        left: origPos.left
    }).animate(properties, options);
}

function jshow_btnEject() {
    $("#btnEject").fadeIn(500);
}

function jHide_btnEject() {
    $("#btnEject").fadeOut(500);
}

