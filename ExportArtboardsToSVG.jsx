/*
KNOWN LIMITATIONS:
  - Poor support for transparency masks

WORKS FOR:
  - Symbols
  - Clipping masks
*/


function slugify(str) {
  var slug = str
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return slug;
}

function findArtboardNameLayer(doc) {
  for (var i = 0; i < doc.layers.length; i++) {
    if (doc.layers[i].name === "artboard_name") {
      return doc.layers[i];
    }
  }
  return null;
}

function getRawArtboardNames(doc) {
  // Returns an array, one entry per artboard, of the trimmed text contents
  // found on the "artboard_name" layer whose bounds center falls inside
  // that artboard, or null if no matching/non-empty text frame was found.
  var rawNames = [];
  var count = doc.artboards.length;
  for (var i = 0; i < count; i++) {
    rawNames.push(null);
  }

  var nameLayer = findArtboardNameLayer(doc);
  if (!nameLayer) {
    return rawNames;
  }

  for (var i = 0; i < count; i++) {
    var rect = doc.artboards[i].artboardRect; // [left, top, right, bottom]
    for (var j = 0; j < nameLayer.textFrames.length; j++) {
      var tf = nameLayer.textFrames[j];
      var bounds = tf.geometricBounds; // [left, top, right, bottom]
      var centerX = (bounds[0] + bounds[2]) / 2;
      var centerY = (bounds[1] + bounds[3]) / 2;

      if (
        centerX >= rect[0] &&
        centerX <= rect[2] &&
        centerY <= rect[1] &&
        centerY >= rect[3]
      ) {
        var contents = tf.contents.replace(/^\s+|\s+$/g, "");
        if (contents.length > 0) {
          rawNames[i] = contents;
        }
        break;
      }
    }
  }

  return rawNames;
}

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

  var options = new ExportOptionsSVG();
  options.embedRasterImages = true;
  options.coordinatePrecision = 4;
  options.preserveEditability = false;
  options.fontType = SVGFontType.OUTLINEFONT; 
  options.cssProperties = SVGCSSPropertyLocation.STYLEATTRIBUTES;
  options.documentEncoding = SVGDocumentEncoding.UTF8;
  options.includeFileInfo = false;

  var count = doc.artboards.length;

  // Read the artboard_name layer (if any) and slugify each name we find,
  // up front, so we can detect duplicates and abort before exporting
  // anything.
  var rawNames = getRawArtboardNames(doc);
  var fileNames = [];
  var slugIndexes = {};

  for (var i = 0; i < count; i++) {
    var slug = rawNames[i] ? slugify(rawNames[i]) : "";

    if (slug.length > 0) {
      if (slugIndexes.hasOwnProperty(slug)) {
        alert(
          'Duplicate artboard name "' +
            slug +
            '" found on artboards ' +
            (slugIndexes[slug] + 1) +
            " and " +
            (i + 1) +
            ".\n\nAborting without exporting any files.",
        );
        return;
      }
      slugIndexes[slug] = i;
      fileNames.push(slug + ".svg");
    } else {
      // CS4 artboards have no .name property (added in CS5), so we number
      // the files by their artboard index instead when no name was found.
      fileNames.push(baseName + "-" + (i + 1) + ".svg");
    }
  }

  // Save a reference to the original .ai file so we can restore it after
  // export. exportFile with SVG performs a "Save As" internally in many
  // Illustrator versions, which changes the document's saved path to the
  // last exported SVG. Saving back to the original path prevents that SVG
  // content from overwriting the .ai file.
  var originalFile = doc.fullName;
  var aiSaveOptions = new IllustratorSaveOptions();

  for (var i = 0; i < count; i++) {
    // Make this artboard the active one before exporting.
    doc.artboards.setActiveArtboardIndex(i);

    var outFile = new File(exportFolder + "/" + fileNames[i]);

    // exportFile overwrites an existing file of the same name.
    doc.exportFile(outFile, ExportType.SVG, options);
  }

  // Restore the document as the original .ai file so Illustrator doesn't
  // leave it in an SVG-saved state.
  doc.saveAs(originalFile, aiSaveOptions);

  alert(
    "Done. Exported " +
      count +
      " artboard(s) to:\n" +
      exportFolder.fsName +
      "\n\nFiles:\n" +
      fileNames.join("\n"),
  );
})();
