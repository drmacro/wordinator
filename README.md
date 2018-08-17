# The Wordinator

Version 0.5.0 

Generate high-quality Microsoft Word DOCX files using a simplified XML format (simple word processing XML).

Simple Word Processing XML (SWPX) makes it relatively easy to transform structured content into DOCX (or other similar word processing formats).

The Wordinator uses the Apache POI library [https://poi.apache.org/] to generate DOCX files from SWPX XML. It uses Saxon [http://saxonica.com/] for XSLT transformations when using the built-in XSLT support.

This approach provides a two-stage X-to-DOCX conversion process, where the first stage is a transform from whatever your input is into one or more SWPX documents and the second stage generates DOCX files from the SWPX files. You can think of the SWPX XML as a very abstract API to the DOCX format.

The Wordinator Java code can run an XSLT to generate the SWPX dynamically from any XML input or you can generate the SWPX documents separately using whatever means you choose and then process those into DOCX files. Word style definitions are managed using a normal Word template (DOTX) file that you create and manage normally.

The Wordinator is designed for batch or on-demand generation of DOCX files.

The Wordinator requires Java 8 or newer (because POI 4 requires it).

The Wordinator provides a generic HTML5-to-DOCX transform that can easily be adapted to your specific HTML or other XML format. 

The main challenges are managing white space within text runs and mapping source elements to the appropriate paragraph and character styles. The XSLT has been designed to make the element-to-style mapping as easy as possible by using a separate XSLT mode to generate the style names for elements. This mode uses an XSLT 3 map to map HTML class values to Word style names and paragraph and run-level formatting controls (e.g., a @class token of 'bold' will result in bold runs). This makes configuring the mapping about as easy as it can be. If a simple class-to-style mapping is insufficient you can use normal XSLT templates to map elements in context to styles.

You can use your own XSLT transform to generate SWPX files from any XML (or JSON source for that matter). You may find it easier to generate HTML and then use that as input to the Wordinator.

If you need to go from Word documents back to XML, you may find the DITA for Publishers Word-to-DITA framework useful ([https://github.com/dita4publishers/org.dita4publishers.word2dita]). This packaged as a DITA Open Toolkit plugin but is really a general-purpose XML-to-DOCX framework. It does not depend on the DITA Open Toolkit in any way. While it is designed to generate DITA XML it can be adapted to produce any XML format, either directly or through a DITA-to-X transform applied 

## Word feature support

The Wordinator supports generation of documents with the following Word features:

* Paragraphs and runs with specific styles
* Footnotes and end notes
* Tables with spans
* Embedded graphics
* Running heads and feet
* Bookmarks
* Hyperlinks
* Multiple sections (full support pending)

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

## Building

This is a Maven project.

NOTE: As of Aug 2018 this code relies on the development version of Apache POI 4.0.0. This means that you'll need to clone or fork the Apache POI sources (e.g., [https://github.com/apache/poi]) and build the POI jars directly following the POI project build instructions, which may be as easy as running the "mvn-install" Ant task.

NOTE: POI 4.0.0 and this project require at least Java 8.

Maven dependency:

```
<dependency>
  <groupId>org.wordinator</groupId>
  <artifactId>wordinator</artifactId>
  <version>0.5.0</version>
</dependency>
```
