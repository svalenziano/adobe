/**********************************************************
 
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

// replace pdfSaveOpts.pDFPreset  = '1.4'; with your distiller preset

// uncomment to suppress Illustrator warning dialogs
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;



var destFolder, sourceFolder, files, fileType, sourceDoc, targetFile, pdfSaveOpts;
 
// Select the source folder.
sourceFolder = Folder.selectDialog( 'Select the folder with Illustrator .ai files you want to convert to PDF');

 
// If a valid folder is selected
if ( sourceFolder != null )
{
    files = new Array();
    fileType = "*.ai"; //prompt( 'Select type of Illustrator files to you want to process. Eg: *.ai', ' ' );
 
    // Get all files matching the pattern
    files = sourceFolder.getFiles( fileType );
 
    if ( files.length > 0 )
    {
        // Get the destination to save the files
        destFolder = Folder.selectDialog( 'Select the folder where you want to save the converted PDF files.', '~' );
        
        logPath = destFolder+'/log.txt';
        var logFile = File(logPath);
        logFile.open('a'); // Open in append mode
        logFile.writeln('This log was created by SV_AI2PDF.jsx (Adobe Illustrator)');
        // Write the current date and time
        logFile.writeln('Timestamp: ' + getCurrentTimestamp());

        for ( i = 0; i < files.length; i++ )
        {
            sourceDoc = app.open(files[i]); // returns the document object
 
            embed_links(logFile);  // Comment this function out if you wish NOT to embed links

            // Call function getNewName to get the name and file to save the pdf
            targetFile = getNewName();
            
            // Call function getPDFOptions get the PDFSaveOptions for the files
            pdfSaveOpts = getPDFOptions();
 
            // Save as pdf
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

/* 
Function to log a message to the text file
    log = log object creaed with File()
    message = string, will be written to file
*/
function logMessage(logObject, message) { 
    logObject.writeln(message);
}


// Helper function to get the current timestamp
function getCurrentTimestamp() {
    var now = new Date();
    var timestamp = now.toLocaleString(); // Format the timestamp as a string
    return timestamp;
}


function embed_links(logObject) {
    // Get all linked items in the document
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