# The Wordinator

Version 0.1.0 

Generate high-quality Microsoft Word DOCX files using a simplified XML format (simple word processing XML).

Simple Word Processing XML (SWPX) makes it relatively easy to transform structured content into DOCX (or other similar word processing formats).

The Wordinator uses the Apache POI library [https://poi.apache.org/] to generate DOCX files from SWPX XML. It uses Saxon [http://saxonica.com/] for XSLT transformations when using the built-in XSLT support.

This approach provides a two-stage X-to-DOCX conversion process, where the first stage is a transform from whatever your input is into one or more SWPX documents and the second stage generates DOCX files from the SWPX files. You can think of the SWPX XML as a very abstract API to the DOCX format.

The Wordinator Java code can run an XSLT to generate the SWPX dynamically from any XML input or you can generate the SWPX documents separately using whatever means you choose and then process those into DOCX files. Word style definitions are managed using a normal Word tempalte (DOTX) file that you create and manage normally.

The Wordinator is designed for batch or on-demand generation of DOCX files.

The Wordinator requires Java 8 or newer (because POI 4 requires it).

The Wordinator provides a generic HTML5-to-DOCX transform that can easily be adapted to your specific HTML or other XML format. The main challenges are managing white space within text runs and mapping source elements to the appropriate paragraph and character styles. The XSLT has been designed to make the element-to-style mapping as easy as possible by using a separate XSLT mode to generate the style names for elements. This mode uses an XSLT 3 map to map HTML class values to Word style names. This makes configuring the mapping about as easy as it can be. If a simple class-to-style mapping is insufficient you can use normal XSLT templates to map elements in context to styles.

## Word feature support

The Wordinator supports generation of documents with the following Word features:

* Paragraphs and runs with specific styles
* Footnotes and end notes
* Tables with spans
* Embedded graphics
* Running heads and feet
* Bookmarks
* Hyperlinks
* Multiple sections

*NOTE:* Use of direct formatting is not supported. All formatting must be defined through the use of templates,

## Getting Started

TBD

## Managing Word Styles

The Wordinator requires a Word template document (DOTX) that defines the styles available in the generated Word document.

To create and manage styles use this general procedure:

1. Create a Word document with the styles you need. For every style, whether built-in or custom, create at least one object (paragraph, character run, table, etc.) that uses the style. 
2. Save the document as a DOTX (Word template document). This will be the template you provide to the Wordinator.
3. To add or modify styles, create a new document from the DOTX file. Going forward you will use this new file to create new styles or modify existing styles.
4. When you create or modify styles in the document, be sure to check the "Add to template" check mark on the style dialog. This will cause the template document to be updated with the new style information when you save the document you are editing.

### Using the Style Organizer

If you forget to do "Add to template" or you want to copy styles from an existing Word document, you can use the style organizer.

To get to the style organizer:
1. Select Tools->Templates and Add-ins to bring up the Tools and Add-ins dialog
  The dialog shows the template associated with the document. If your template is not attached, use the Attach button to attach it.
  Make sure the "Automatically update document styles" check box is checked.
2. Click the "Organizer" button to open the Organizer dialog. Select the "Styles" tab
3. The right side of the Organizer dialog shows the template document to which you will copy styles. It probably shows the default template. If so, click "Close file" and then click "Open file" and select your DOTX file.
4. Use the Organizer dialog to copy styles from the left side to the right side.
5. Click "Close" to save your changes to the template document.

## Support, New Feature Development, and Contributing

The Wordinator project is supported primarily by paying clients who fund development of the features they need. Initial development was funded by Municode [http://municode.com].

Please use this project's issue tracker to report bugs or request new features. 

I (Eliot Kimber) will attempt to fix bugs as quickly as possible.

For new features, it is unlikely that I will be able to implement them outside of a paid engagement, but if it's something generally usable or something one of my clients needs I may be able to implement it.

If you would like to contribute new features, I welcome all contributions. Use normal GitHub pull requests to submit your contributions. If you'd like to be more heavily involved or even take over primarily development, please contact me directly.


