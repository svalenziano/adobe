// ExportArtboardsToSVG.jsx
// Proof of concept for Adobe Illustrator CS4.
// Exports every artboard to SVG into an "export" folder that sits in the
// same directory as the .ai file. Existing files are overwritten.
//
// To run: File > Scripts > Other Script...  (or drag onto Illustrator)

(function () {
  if (app.documents.length === 0) {
    alert("Open your .ai document first, then run the script.");
    return;
  }

  var doc = app.activeDocument;

  // We need to know where the .ai file lives on disk.
  var docFolder;
  try {
    docFolder = doc.path; // Folder object for a saved document
  } catch (e) {
    docFolder = null;
  }
  if (!docFolder) {
    alert(
      "Please save the .ai file to disk first, so I know where to create the export folder.",
    );
    return;
  }

  // Create (or reuse) the "export" folder next to the .ai file.
  var exportFolder = new Folder(docFolder + "/export");
  if (!exportFolder.exists) {
    exportFolder.create();
  }

  // Base name = the .ai filename without its extension.
  var baseName = doc.name.replace(/\.[^\.]+$/, "");

  // Default SVG export options. Keeping it bare for the proof of concept.
  var options = new ExportOptionsSVG();

  var count = doc.artboards.length;

  for (var i = 0; i < count; i++) {
    // Make this artboard the active one before exporting.
    doc.artboards.setActiveArtboardIndex(i);

    // CS4 artboards have no .name property (added in CS5), so we number
    // the files by their artboard index instead.
    var fileName = baseName + "-" + (i + 1) + ".svg";

    var outFile = new File(exportFolder + "/" + fileName);

    // exportFile overwrites an existing file of the same name.
    doc.exportFile(outFile, ExportType.SVG, options);
  }

  alert("Done. Exported " + count + " artboard(s) to:\n" + exportFolder.fsName);
})();
