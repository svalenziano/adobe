function closeAllDocsWithoutSaving() {
    // Suppress Illustrator's warning dialogs
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
    
    while (app.documents.length > 0) {
        app.documents[0].close(SaveOptions.DONOTSAVECHANGES);
    }

    // Restore the default user interaction level
    app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
}

function yes_or_no(message){
    var userResponse = confirm(message);
    if (userResponse) {
        // User clicked 'Yes' (OK)
        return true
    } else {
        // User clicked 'No' (Cancel)
        return false
    }
}

/******************************************* 
 MAIN LOOP
*******************************************/

var areYouSure = yes_or_no("CLOSE ALL DOCS WITHOUT SAVING?")

if (areYouSure == true) {
    closeAllDocsWithoutSaving();
}
