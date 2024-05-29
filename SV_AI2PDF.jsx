/**********************************************************

WIP:
PROVIDING THE ABILITY TO WORK WITH EITHER OPENED FILES *OR* A FOLDER FULL O FILES, SEE #LOA


DESCRIPTION
 
This sample gets files specified by the user from the 
selected folder and batch processes them and saves them 
as PDFs.

Errors are logged to Illustrator's JavaScript Console (Ctrl+Alt+F12 on Windows)

Edits by SV:
 - Modify PDF save options
 - Embed Links function

Edit by Swasher:
 - only .ai files processed
 - filenames not changes
 - files saved in same folder as the input files
 - pdfsetting got from adobe setting ('1.4' in this file)
 - perform black overprint
 - new function for renaming .ai to .pdf (old function has bug with filenames which contain dots)

USAGE:
 - put script in Illustrator script folder (something like
   C:\Program Files\Adobe\Adobe Illustrator CC 2014\Presets\en_GB\Scripts\AI2PDF.jsx)
 - in illustrator menu choose File->Scripts->AI2PDF and click it.
 - select appropriate folder with ai files
 - enjoy

CREDITS:
Forked from Swasher

**********************************************************/

// uncomment to suppress Illustrator warning dialogs
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

var destFolder, files, sourceDoc, targetFile, pdfSaveOpts;

/**********************************************************
SETTINGS!
**********************************************************/
var useOpenFiles = yes_or_no("Use open documents?  They will be saved as PDFs and closed.")
var embed = true;
var closeWhenDone = false

/**********************************************************
MAIN LOOP
**********************************************************/
if (useOpenFiles == false) {
    files = getFilesFromFolder();

    if ( files.length > 0 )
    {
        // Get the destination to save the files
        destFolder = Folder.selectDialog( 'DESTINATION FOLDER', '~' );
        
        logFile = createLog(destFolder+'/log.txt')

        for ( i = 0; i < files.length; i++ )
        {
            sourceDoc = app.open(files[i]); // returns the document object

            if (embed == true) {
                embed_links(logFile, sourceDoc);  // Comment this function out if you wish NOT to embed links
            }

            // Call function getNewName to get the name and file to save the pdf
            targetFile = getNewName();
            
            // Call function getPDFOptions get the PDFSaveOptions for the files
            pdfSaveOpts = getPDFOptions();

            // Save as pdf
            logMessage(logObject, sourceDoc.name)   
            sourceDoc.saveAs(targetFile,pdfSaveOpts);

            sourceDoc.close();
        }
        logFile.close();
        alert( 'Files are saved as PDF in ' + destFolder );
    }
    else
    {
        alert( 'No matching files found' );
    }
}

else {
    if (app.documents.length > 0 ) {
        
        // Get the folder to save the files into
        var destFolder = null;
        destFolder = Folder.selectDialog( 'DESTINATION FOLDER', '~' );
        
        if (destFolder != null) {
            var pdfSaveOptions, i, sourceDoc, targetFile;	
            
            // closeWhenDone = yes_or_no("Close PDFs when done?")  // Not working :(

            logObject = createLog(destFolder + '/log.txt')

            
            // Get the PDF options to be used
            pdfSaveOptions = getPDFOptions();

            var docCount = app.documents.length;
                                    
            for ( i = 0; i < docCount; i++ ) {
                sourceDoc = app.documents[i]; // returns the document object
                logMessage(logObject, sourceDoc.name)

                if (embed == true) {
                    embed_links(logObject, sourceDoc);  // Comment this function out if you wish NOT to embed links
                }

                // Get the file to save the document as pdf into
                targetFile = this.getTargetFile(sourceDoc.name, '.pdf', destFolder);
                
                // Save as pdf
                sourceDoc.saveAs( targetFile, pdfSaveOptions );
                
                if (closeWhenDone == true){
                    // Close so you don't have a bunch of open (unsaved) documents to deal with
                    sourceDoc.close();
                }
            }
            logObject.close();
            alert( 'Files are saved as PDF in ' + destFolder );
        }
    }
    else{
        throw new Error('There are no documents open!');
    }
}


/**********************************************************
FUNCTIONS
**********************************************************/

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


/*
This function is from the sample script that ships with Illustrator
*/
function getTargetFile(docName, ext, destFolder) {
	var newName = "";

	// if name has no dot (and hence no extension),
	// just append the extension
	if (docName.indexOf('.') < 0) {
		newName = docName + ext;
	} else {
		var dot = docName.lastIndexOf('.');
		newName += docName.substring(0, dot);
		newName += ext;
	}
	
	// Create the file object to save to
	var myFile = new File( destFolder + '/' + newName );
	
	// Preflight access rights
	if (myFile.open("w")) {
		myFile.close();
	}
	else {
		throw new Error('Access is denied');
	}
	return myFile;
}

/* 
Function to log a message to the text file
    log = log object creaed with File()
    message = string, will be written to file
*/

function getOpenFiles(){
    var files
    // #LOA, uh oh, realizing I'm not really sure what I'm doing here!


}

function getFilesFromFolder() {
    // Select the source folder.
    var sourceFolder;
    var files;
    var fileType;
    sourceFolder = Folder.selectDialog( 'Select the folder with Illustrator .ai files you want to convert to PDF');
    
    // If a valid folder is selected
    if ( sourceFolder != null )
    {
        files = new Array();
        fileType = "*.ai"; //prompt( 'Select type of Illustrator files to you want to process. Eg: *.ai', ' ' );
    
        // Get all files matching the pattern
        files = sourceFolder.getFiles( fileType );
    }
    return files
}

function createLog(logPath) {
    var logFile = File(logPath);
    logFile.open('a'); // Open in append mode
    logFile.writeln('');
    logFile.writeln('');
    logFile.writeln('This log was created by SV_AI2PDF.jsx (Adobe Illustrator)');
    // Write the current date and time
    logFile.writeln(getCurrentTimestamp());
    return logFile
}

function logMessage(logObject, message) { 
    logObject.writeln(message);
}


// Helper function to get the current timestamp
function getCurrentTimestamp() {
    var now = new Date();
    var timestamp = now.toLocaleString(); // Format the timestamp as a string
    return timestamp;
}

// Embed links and log to file
function embed_links(logObject,sourceDoc) {
    // Get all linked items in the document
    var link_count = sourceDoc.placedItems.length;
    if (link_count > 0) {
        // logMessage(logObject, "Embedding " + link_count + " links...");
        var linkedItems = sourceDoc.placedItems;
    
        // Loop through each linked item and embed it
        for (var i = 0; i < linkedItems.length; i++) {
            var linkedItem = linkedItems[i];
            try {
                linkedItem.embed();
                // If successful, log a message
                logMessage(logObject, "Embedded: " + linkedItem.file.name);
            } catch (e) {
                // If embedding fails, log an error message
                logMessage(logObject, "Error embedding: " + linkedItem.file.name);
            }
        }
    }
}



/*********************************************************
 
getNewName: Function to get the new file name. The primary
name is the same as the source file.
 
**********************************************************/
 
function getNewName()
{
    var ext, docName, newName, saveInFile;
    docName = sourceDoc.name;
    ext = '.pdf'; // new extension for pdf file

    newName = docName.substr(0, docName.lastIndexOf(".")) + ext;

    // Create a file object to save the pdf
    saveInFile = new File( destFolder + '/' + newName );
 
    return saveInFile;
}
 
/*********************************************************
 
getPDFOptions: Function to set the PDF saving options of the 
files using the PDFSaveOptions object.
 
**********************************************************/
 
function getPDFOptions() {
    // Create the PDFSaveOptions object to set the PDF options
    var pdfSaveOpts = new PDFSaveOptions();

    // Set PDF options
    pdfSaveOpts.acrobatLayers = true;
    pdfSaveOpts.colorDownsampling = 0;
    pdfSaveOpts.compatibility = PDFCompatibility.ACROBAT8;
    pdfSaveOpts.compressArt = false;
    pdfSaveOpts.grayscaleDownsampling = 0;
    pdfSaveOpts.monochromeDownsampling = 0;
    pdfSaveOpts.preserveEditability = true;
    pdfSaveOpts.viewAfterSaving = false;

    return pdfSaveOpts;
}



/*********************************************************

performBlackOverprint: Function to set all black objects as overprinted.
 Source at http://www.typomedia.org/adobe/illustrator/black-overprint/

**********************************************************/

function performBlackOverprint(doc)
{

    // var doc = sourceDoc;
    var k;

    if(doc.documentColorSpace == DocumentColorSpace.RGB) {
        alert("The working color space is RGB. Change the document color mode to CMYK.",
            "CMYK working space required!");
        }

    else {

        // Skip Blank text frames
        for ( k = 0 ; k < doc.textFrames.length ; k++) {
            if (doc.textFrames[k].contents != "")	{
                txt=doc.textFrames[k];

                // Black text fill to overprint
                if((txt.textRange.characterAttributes.overprintFill == false
                    && txt.textRange.characterAttributes.fillColor.cyan == 0
                    && txt.textRange.characterAttributes.fillColor.magenta == 0
                    && txt.textRange.characterAttributes.fillColor.yellow == 0
                    && txt.textRange.characterAttributes.fillColor.black == 100
                    || txt.textRange.characterAttributes.fillColor.gray == 100
                    )){
                txt.textRange.characterAttributes.overprintFill=true;
                }

                // Black text contour to overprint
                if((txt.textRange.characterAttributes.overprintStroke == false
                    && txt.textRange.characterAttributes.strokeColor.cyan == 0
                    && txt.textRange.characterAttributes.strokeColor.magenta == 0
                    && txt.textRange.characterAttributes.strokeColor.yellow == 0
                    && txt.textRange.characterAttributes.strokeColor.black == 100
                    || txt.textRange.characterAttributes.strokeColor.gray == 100
                    )){
                txt.textRange.characterAttributes.overprintStroke=true;
                }

            }
        }

        // Page properties
        for( k = 0 ; k < doc.pathItems.length; k++){
            obj = doc.pathItems[k];

            // Black fills to overprint
            if((obj.fillOverprint == false
                && obj.fillColor.cyan == 0
                && obj.fillColor.magenta == 0
                && obj.fillColor.yellow == 0
                && obj.fillColor.black == 100
                || obj.fillColor.gray == 100
                )){
            obj.fillOverprint=true;
            }

            // Black contours on overprinting
            if((obj.strokeOverprint == false
                && obj.strokeColor.cyan == 0
                && obj.strokeColor.magenta == 0
                && obj.strokeColor.yellow == 0
                && obj.strokeColor.black == 100
                || obj.strokeColor.gray == 100
                )){
            obj.strokeOverprint=true;
            }
        }

        // alert("Black outlines and fillings were set to overprint.", "The changes have been made!");
    }
}