package org.wordinator.xml2docx;

import java.io.File;
import java.io.FileInputStream;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.apache.poi.ooxml.POIXMLProperties.CustomProperties;
import org.apache.poi.ooxml.POIXMLProperties.ExtendedProperties;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.IRunElement;
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFComment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFFootnote;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFNum;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.junit.Test;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFldChar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColumns;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.wordinator.xml2docx.generator.DocxConstants;
import org.wordinator.xml2docx.generator.DocxGenerationException;
import org.wordinator.xml2docx.generator.DocxGenerator;

import junit.framework.TestCase;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.apache.poi.util.Units;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STAlignH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STAlignV;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromV;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STWrapText;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

public class TestDocxGenerator extends TestCase {

  public static final String DOTX_TEMPLATE_PATH = "docx/Test_Template.dotx";
  
  static final SimpleDateFormat isoDateFormatter = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ssX", Locale.ENGLISH);

  @Test
  public void testMakeDocx() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-test-01.swpx", "out/testMakeDocx.docx");
    Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
    XWPFParagraph p = iterator.next();
    assertNotNull("Expected a paragraph", p);
    assertEquals("Heading 1 Text", p.getText());
    System.out.println("Paragraph text='" + p.getText() + "'");
    boolean foundToc = false;
    while (iterator.hasNext()) {
      p = iterator.next();
      // Issue 42: Get style ID
      String styleId = p.getStyle();
      // Issue 16: Verify scaling of intrinsic dimensions:
      if (p.getText().startsWith("[Image 1]")) {
        XWPFRun run = p.getRuns().get(1); // Second run should contain the picture
        assertNotNull("Extected a second run", run);
        XWPFPicture picture = run.getEmbeddedPictures().get(0);
        assertNotNull("Expected a picture", picture);
        assertEquals("Expected width of 100", picture.getWidth(), 100.0);
        assertEquals("Expected height (depth) of 50", picture.getDepth(), 50.0);
      }
      if (p.getText().startsWith("[Image 2]")) {
        XWPFRun run = p.getRuns().get(1); // Second run should contain the picture
        assertNotNull("Extected a second run", run);
        XWPFPicture picture = run.getEmbeddedPictures().get(0);
        assertNotNull("Expected a picture", picture);
        assertEquals("Expected width of 100", picture.getWidth(), 100.0);
        assertEquals("Expected height (depth) of 100", picture.getDepth(), 100.0);
      }
      if (p.getText().startsWith("[Image 3]")) {
        XWPFRun run = p.getRuns().get(1); // Second run should contain the picture
        assertNotNull("Extected a second run", run);
        XWPFPicture picture = run.getEmbeddedPictures().get(0);
        assertNotNull("Expected a picture", picture);
        assertEquals("Expected width of 50", picture.getWidth(), 50.0);
        assertEquals("Expected height (depth) of 50", picture.getDepth(), 50.0);
      }
      if ("TOC1".equals(styleId)) {
        foundToc = true;
      }
    }

    assertTrue("Did not find expected table of contents paragraphs", foundToc);
  }

  @Test
  public void testMakeDocxWithSections() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-test-02.swpx", "out/output-02.docx");
    Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
    XWPFParagraph p = iterator.next();
    assertNotNull("Expected a paragraph", p);
    assertEquals("Document With Sections", p.getText());
    System.out.println("Paragraph text='" + p.getText() + "'");

    boolean found = false;
    for (XWPFParagraph para : doc.getParagraphs()) {
      String text = para.getParagraphText();
      // System.out.printn("+ [DEBUG] text=\"" + text + "\"");
      if ("Document With Sections".equals(text)) {
        found = true;
        break;
      }
    }
    assertTrue("Did not find expected start of first section", found);

    CTSectPr docSectPr = doc.getDocument().getBody().getSectPr();
    assertNotNull("Expected to find a docSectPr element", docSectPr);
    // The last section's headers and footers replace the document-level ones
    // of the same type on the body sectPr (issue #172), so one of each type at most.
    assertEquals("Expected 3 headers", 3, docSectPr.getHeaderReferenceList().size());
    assertEquals("Expected 3 footers", 3, docSectPr.getFooterReferenceList().size());

    // Document-level headers and footers:
    XWPFHeaderFooterPolicy hfPolicy = doc.getHeaderFooterPolicy();

    // Headers:
    XWPFHeader header = hfPolicy.getDefaultHeader();
    List<IBodyElement> bodyElems = header.getBodyElements();
    assertEquals("Expected 1 paragraph", 1, bodyElems.size());
    header = hfPolicy.getEvenPageHeader();
    bodyElems = header.getBodyElements();
    assertEquals("Expected 1 paragraph", 1, bodyElems.size());
    header = hfPolicy.getFirstPageHeader();
    bodyElems = header.getBodyElements();
    assertEquals("Expected 1 paragraph", 1, bodyElems.size());

    // Footers:
    XWPFFooter footer = hfPolicy.getDefaultFooter();
    bodyElems = footer.getBodyElements();
    assertEquals("Expected 1 paragraph", 1, bodyElems.size());
    footer = hfPolicy.getEvenPageFooter();
    bodyElems = header.getBodyElements();
    assertEquals("Expected 1 paragraph", 1, bodyElems.size());
    footer = hfPolicy.getFirstPageFooter();
    bodyElems = header.getBodyElements();
    assertEquals("Expected 1 paragraph", 1, bodyElems.size());

    // Section headers and footers:

    boolean foundHeadersOrFooters = false;
    Iterator<IBodyElement> iter = doc.getBodyElementsIterator();
    do {
      IBodyElement e = iter.next();
      if (e instanceof XWPFParagraph) {
        p = (XWPFParagraph)e;
        if (p.getCTP().isSetPPr()) {
          CTSectPr sectPr = p.getCTP().getPPr().getSectPr();
          if (sectPr != null) {
            assertTrue("Expected no more than 3 headers for paragraph, found " + sectPr.getHeaderReferenceList().size(),
                       sectPr.getHeaderReferenceList().size() <= 3);
            assertTrue("Expected no more than 3 footers for paragraph, found " + sectPr.getFooterReferenceList().size(),
                       sectPr.getFooterReferenceList().size() <= 3);
            foundHeadersOrFooters = foundHeadersOrFooters ||
              sectPr.getHeaderReferenceList().size() > 0  ||
              sectPr.getFooterReferenceList().size() > 0;
          }
        }
      }
    } while(iter.hasNext());
    assertTrue("No section headers or footers", foundHeadersOrFooters);
  }

  /**
   * Issue #172: Headers and footers of the last section were added to the body-level
   * section properties alongside the document-level ones, producing duplicate
   * references of the same type and doubled headers and footers in Word.
   */
  @Test
  public void testNoDuplicateHeaderFooterReferencesInLastSection() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-test-02.swpx", "out/output-issue-172.docx");
    CTSectPr docSectPr = doc.getDocument().getBody().getSectPr();
    assertNotNull("Expected a body-level sectPr", docSectPr);
    java.util.Set<String> types = new java.util.HashSet<String>();
    for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtrRef ref : docSectPr.getHeaderReferenceList()) {
      assertTrue("Duplicate header reference of type " + ref.getType(), types.add(ref.getType().toString()));
    }
    types.clear();
    for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtrRef ref : docSectPr.getFooterReferenceList()) {
      assertTrue("Duplicate footer reference of type " + ref.getType(), types.add(ref.getType().toString()));
    }
  }

  /**
   * Issue #172: The type of the last section (how it starts) must be carried over
   * to the body-level section properties when they replace the last paragraph's.
   */
  @Test
  public void testLastSectionTypeCopiedToBodySectPr() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-172-last-section-type.swpx", "out/output-issue-172-type.docx");
    CTSectPr docSectPr = doc.getDocument().getBody().getSectPr();
    assertNotNull("Expected a body-level sectPr", docSectPr);
    assertTrue("Expected type to be set on body-level sectPr", docSectPr.isSetType());
    assertEquals("Expected last section type", "oddPage", docSectPr.getType().getVal().toString());
  }

  @Test
  public void testCopyNumberingDefinitions() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-test-03.swpx", "out/output-03.docx");
    Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
    XWPFParagraph p = iterator.next();
    assertNotNull("Expected a paragraph", p);
    assertEquals("Test of List Formatting", p.getText());
    System.out.println("Paragraph text='" + p.getText() + "'");
    XWPFNumbering numbering = doc.getNumbering();
    assertNotNull("No numbering", numbering);

    XWPFAbstractNum abstractNumber = numbering.getAbstractNum(BigInteger.valueOf(9));
    assertNotNull("No abstract number '9'", abstractNumber);

    XWPFNum num = numbering.getNum(BigInteger.valueOf(9));
    assertNotNull("No num '9'", num);
  }

  @Test
  public void testFieldGeneration() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-47.swpx", "out/output-issue-47.docx");
    Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
    XWPFParagraph p = iterator.next();
    assertNotNull("Expected a paragraph", p);
    assertEquals("Issue 47: Complex Fields", p.getText());
    p = iterator.next();
    assertNotNull("Expected a paragraph", p);
    assertTrue("Didn't find lead-in to field", p.getText().startsWith("Complex field"));
    List<XWPFRun> runs = p.getRuns();
    boolean foundField = false;
    boolean foundStart = false;
    boolean foundEnd = false;
    boolean foundSeparator = false;
    String instructionText = null;
    for (XWPFRun run : runs) {
      // Check for field here
      CTR r = run.getCTR();
      List<CTFldChar> fldChars = r.getFldCharList();
      if (fldChars != null && fldChars.size() > 0) {
        if (fldChars.get(0).getFldCharType() == STFldCharType.BEGIN) {
          foundStart = true;
        }
        if (fldChars.get(0).getFldCharType() == STFldCharType.END) {
          foundEnd = true;
        }
        if (fldChars.get(0).getFldCharType() == STFldCharType.SEPARATE) {
          foundSeparator = true;
        }
      }
      List<CTText> instructions = r.getInstrTextList();
      if (instructions != null && instructions.size() > 0) {
        instructionText = instructions.get(0).getStringValue();
      }
    }
    foundField = foundStart && foundEnd;
    assertTrue("Did not find expected field", foundField);
    assertTrue("Expected to find a separator", foundSeparator);
    assertEquals("Instruction text did not match", "DATE \\@ \"dddd, MMMM dd, yyyy HH:mm:ss\"", instructionText);
  }

  @Test
  public void testFootnoteGeneration() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-29.swpx", "out/output-issue-47.docx");
    Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
    XWPFParagraph p = iterator.next();
    assertNotNull("Expected a paragraph", p);
    assertEquals("Issue #29: Test of Literal Footnote Callouts", p.getText());
    // Normal footnote with generated ref
    p = iterator.next();
    assertEquals(" [1: This is a normal footnote. It should have a callout of 1] ", p.getFootnoteText());
    // Custom footnote with literal callout with same value for ref and footnote:
    p = iterator.next();
    String fnText = p.getFootnoteText();
    assertEquals(" [2: FN-1This is a custom footnote. It specifies a literal callout of \"FN-1\".] ", fnText);
    // Custom footnote with literal callout with different values for ref and footnote:
    p = iterator.next();
    fnText = p.getFootnoteText();
    assertEquals(" [3: FN-2This is a custom footnote. It specifies a literal callout of \"FN-2\" and a reference callout of \"Ref-2\".] ", p.getFootnoteText());
  }

  @Test
  public void testEmptyFootnoteCallout() throws Exception {
    // Issue #154: callout="" must be treated as no literal callout.
    XWPFDocument doc = convert("simplewp/simplewpml-issue-154-empty-callout.swpx", "out/issue-154-empty-callout.docx");
    XWPFParagraph p = null;
    for (XWPFParagraph candidate : doc.getParagraphs()) {
      if (candidate.getText().contains("empty callout")) {
        p = candidate;
      }
    }
    assertNotNull("Expected the paragraph with the footnote", p);
    String paraXml = p.getCTP().xmlText();
    assertFalse("Reference should not use a custom mark", paraXml.contains("customMarkFollows"));
    assertFalse("Reference should not contain an empty literal callout", paraXml.contains("<w:t/>"));
    XWPFFootnote footnote = doc.getFootnotes().stream()
        .filter(fn -> fn.getParagraphs().stream().anyMatch(fp -> fp.getText().contains("Footnote text")))
        .findFirst().orElse(null);
    assertNotNull("Expected the footnote", footnote);
    String noteXml = footnote.getCTFtnEdn().xmlText();
    assertTrue("Footnote should keep its generated footnoteRef", noteXml.contains("footnoteRef"));
  }

  @Test
  public void testHyperlinkInFootnote() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-footnote-hyperlink.swpx", "out/footnote-hyperlink.docx");
    XWPFFootnote footnote = doc.getFootnotes().stream()
        .filter(fn -> fn.getParagraphs().stream().anyMatch(fp -> fp.getText().contains("the target")))
        .findFirst().orElse(null);
    assertNotNull("Expected a footnote containing the hyperlink text", footnote);
    String rId = null;
    for (XWPFParagraph fp : footnote.getParagraphs()) {
      for (CTHyperlink link : fp.getCTP().getHyperlinkList()) {
        rId = link.getId();
      }
    }
    assertNotNull("Expected a hyperlink within the footnote", rId);
    // The relationship must live in the footnote part, not the main document part,
    // otherwise Word reports the document as corrupt.
    PackageRelationship rel = footnote.getPart().getPackagePart().getRelationship(rId);
    assertNotNull("Expected relationship " + rId + " in the footnotes part", rel);
    assertEquals("https://example.com/footnote-target", rel.getTargetURI().toString());
    assertEquals(TargetMode.EXTERNAL, rel.getTargetMode());
  }

  @Test
  public void testHyperlinkHandling() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-65.swpx", "out/output-issue-65.docx");
    Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
    XWPFParagraph p = iterator.next();
    assertNotNull("Expected a paragraph", p);
    assertEquals("Issue 65: Data following hyperlink is dropped", p.getText());
    // Hyperlink to a bookmark (internal link):
    p = iterator.next();
    Iterator<IRunElement> runIterator = p.getIRuns().iterator();
    assertTrue("Expected runs", runIterator.hasNext());
    IRunElement run = runIterator.next();
    assertEquals("First run in para before the hyperlink", ((XWPFRun)run).getText(0));
    run = runIterator.next(); // Should be the hyperlink
    assertTrue("Expected a XWPFHyperlinkRun", run instanceof XWPFHyperlinkRun);
    assertTrue("Expected run following the hyperlink", runIterator.hasNext());
    run = runIterator.next(); 
    assertEquals("Run after the hyperlink.", ((XWPFRun)run).getText(0));
    // Now look for external link:
    p = iterator.next();
    assertNotNull("Expected another paragraph after internal hyperlink", p);
    p = iterator.next();
    assertNotNull("Expected another paragraph after internal hyperlink", p);
    p = iterator.next(); // Should be paragraph with hyperlink
    assertNotNull("Expected another paragraph after internal hyperlink", p);
    runIterator = p.getIRuns().iterator();
    assertTrue("Expected runs", runIterator.hasNext());
    run = runIterator.next();
    assertTrue("First run in para before the hyperlink", ((XWPFRun)run).getText(0).startsWith("External hyperlink:"));
    run = runIterator.next(); // Should be the hyperlink
    assertTrue("Expected a XWPFHyperlinkRun", run instanceof XWPFHyperlinkRun);
    XWPFHyperlinkRun linkRun = (XWPFHyperlinkRun)run;
    String rId = linkRun.getHyperlinkId();
    assertNotNull("Expected and rId", rId);
    
  }

  @Test
  public void testTableGeneration() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-30.swpx", "out/output-issue-30.docx");
    Iterator<XWPFTable> iterator = doc.getTablesIterator();

    XWPFTable table;
    table = iterator.next();
    assertNotNull("Did not find any tables", table);

    // Look for table details here.

    // System.out.println("Table rows:");
    /*
      <w:tc>
      <w:tcPr>
      <w:tcW w:w="2880" w:type="dxa"/>
      <w:tcBorders>
      <w:top w:val="single"/>
      <w:left w:val="single"/>
      <w:bottom w:val="single"/>
      <w:right w:val="single"/>
      </w:tcBorders>
      </w:tcPr>
    */
    // int n = 0;
    // First table should have all single borders
    for (XWPFTableRow row : table.getRows()) {
      // System.out.println("Row " + ++n);
      for (XWPFTableCell cell : row.getTableCells()) {
        XmlCursor cursor = cell.getCTTc().newCursor();
        assertTrue("No tcPr element", cursor.toChild(DocxConstants.QNAME_TCPR_ELEM));
        assertTrue("No tcBorders element", cursor.toChild(DocxConstants.QNAME_TCBORDERS_ELEM));
        assertTrue("No top element", cursor.toChild(DocxConstants.QNAME_TOP_ELEM));
        assertEquals("single", cursor.getAttributeText(DocxConstants.QNAME_VAL_ATT));
        assertNull("Did not expect a color attribute", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
        cursor.toNextSibling(); // Left
        assertEquals("single", cursor.getAttributeText(DocxConstants.QNAME_VAL_ATT));
        assertNull("Did not expect a color attribute", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
        cursor.toNextSibling(); // Bottom
        assertEquals("single", cursor.getAttributeText(DocxConstants.QNAME_VAL_ATT));
        assertNull("Did not expect a color attribute", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
        cursor.toNextSibling(); // Right
        assertEquals("single", cursor.getAttributeText(DocxConstants.QNAME_VAL_ATT));
        assertNull("Did not expect a color attribute", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
      }
    }

    // Issue 49: Should have fixed table layout
    XmlCursor cursor = table.getCTTbl().newCursor();
    cursor.push();
    assertTrue("Expected a w:tblPr child", cursor.toChild(DocxConstants.QNAME_TBLPR_ELEM));
    assertTrue("Expected a w:tblLayout child", cursor.toChild(DocxConstants.QNAME_TBLLAYOUT_ELEM));
    assertEquals("expected value 'autofit' for tblLayout", "autofit", cursor.getAttributeText(DocxConstants.QNAME_WTYPE_ATT));
    cursor.pop();

    assertTrue("Expected a second table", iterator.hasNext());
    table = iterator.next();

    for (XWPFTableRow row : table.getRows()) {
      // System.out.println("Row " + ++n);
      for (XWPFTableCell cell : row.getTableCells()) {
        cursor = cell.getCTTc().newCursor();
        assertTrue("No tcPr element", cursor.toChild(DocxConstants.QNAME_TCPR_ELEM));
        assertTrue("No tcBorders element", cursor.toChild(DocxConstants.QNAME_TCBORDERS_ELEM));
        assertTrue("No top element", cursor.toChild(DocxConstants.QNAME_TOP_ELEM));
        assertNotNull("Expected a color attribute", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
        cursor.toNextSibling(); // Left
        assertNotNull("Expected a color attribute", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
        cursor.toNextSibling(); // Bottom
        assertNotNull("Expected a color attribute", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
        cursor.toNextSibling(); // Right
        assertNotNull("Expected a color attribute", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
      }
    }

    assertTrue("Expected a third table", iterator.hasNext());
    table = iterator.next();

    // Issue 49: Should have auto table layout
    cursor = table.getCTTbl().newCursor();
    cursor.push();
    assertTrue("Expected a w:tblPr child", cursor.toChild(DocxConstants.QNAME_TBLPR_ELEM));
    assertTrue("Expected a w:tblLayout child", cursor.toChild(DocxConstants.QNAME_TBLLAYOUT_ELEM));
    assertEquals("expected value 'autofit' for tblLayout", "autofit", cursor.getAttributeText(DocxConstants.QNAME_WTYPE_ATT));
    cursor.pop();

    XWPFTableRow row = table.getRow(0); // Header row.
    XWPFTableCell cell = row.getCell(1); // Center cell
    cursor = cell.getCTTc().newCursor();
    assertTrue(cursor.toChild(DocxConstants.QNAME_TCPR_ELEM));
    assertTrue(cursor.toChild(DocxConstants.QNAME_TCBORDERS_ELEM));
    assertTrue(cursor.toChild(DocxConstants.QNAME_LEFT_ELEM));
    assertEquals("00FF00", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
    cursor.toParent();
    assertTrue(cursor.toChild(DocxConstants.QNAME_RIGHT_ELEM));
    assertEquals("0000FF", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
    cursor.toParent();
    assertTrue(cursor.toChild(DocxConstants.QNAME_TOP_ELEM));
    assertEquals("F0F0F0", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
    cursor.toParent();
    assertTrue(cursor.toChild(DocxConstants.QNAME_BOTTOM_ELEM));
    assertEquals("0F0F0F", cursor.getAttributeText(DocxConstants.QNAME_COLOR_ATT));
  }

  @Test
  public void testPageMarginsDocLevel() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-46-01.swpx", "out/output-issue-46-01.docx");
    CTDocument1 ctDoc = doc.getDocument();
    CTBody body = ctDoc.getBody();
    CTSectPr sectPr = body.getSectPr();
    CTPageMar ctMargins = sectPr.getPgMar();
    assertNotNull("Did not find a CTPageMar object", ctMargins);
    //  <page-margins top="2.0cm" bottom="3.0cm" left="2.5cm" right="3.5cm" footer="1.27cm" header="1.27cm" gutter="0" />
    XmlCursor cursor = ctMargins.newCursor();
    String attVal = cursor.getAttributeText(DocxConstants.QNAME_OOXML_LEFT_ATT);
    assertEquals("1418", attVal);
    attVal = cursor.getAttributeText(DocxConstants.QNAME_OOXML_RIGHT_ATT);
    assertEquals("1985", attVal);
    attVal = cursor.getAttributeText(DocxConstants.QNAME_OOXML_TOP_ATT);
    assertEquals("1134", attVal);
    attVal = cursor.getAttributeText(DocxConstants.QNAME_OOXML_BOTTOM_ATT);
    assertEquals("1701", attVal);
  }

  @Test
  public void testPageMarginsSectionLevel() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-46-02.swpx", "out/output-issue-46-02.docx");
    CTDocument1 ctDoc = doc.getDocument();
    CTBody body = ctDoc.getBody();
    CTSectPr sectPr = body.getSectPr();
    CTPageMar ctMargins = sectPr.getPgMar();
    assertNotNull("Did not find a CTPageMar object", ctMargins);
    //  <page-margins top="2.0cm" bottom="3.0cm" left="2.5cm" right="3.5cm" footer="1.27cm" header="1.27cm" gutter="0" />
    XmlCursor cursor = ctMargins.newCursor();
    String attVal = cursor.getAttributeText(DocxConstants.QNAME_OOXML_LEFT_ATT);
    assertEquals("1418", attVal);
    attVal = cursor.getAttributeText(DocxConstants.QNAME_OOXML_RIGHT_ATT);
    assertEquals("1985", attVal);
    attVal = cursor.getAttributeText(DocxConstants.QNAME_OOXML_TOP_ATT);
    assertEquals("1134", attVal);
    attVal = cursor.getAttributeText(DocxConstants.QNAME_OOXML_BOTTOM_ATT);
    assertEquals("1701", attVal);
  }

  @Test
  public void testPageLayout() throws Exception {
    // Using the issue-46-02.swpx because it happens to have a section with landscape pages.
    XWPFDocument doc = convert("simplewp/simplewpml-issue-46-02.swpx", "out/output-issue-46-02.docx");
    Iterator<IBodyElement> iter = doc.getBodyElementsIterator();
    XWPFParagraph p = null;
    int sectionCounter = 0;
    boolean foundPageLayout = false;
    do {
      IBodyElement e = iter.next();
      if (e instanceof XWPFParagraph) {
        p = (XWPFParagraph)e;
        if (p.getCTP().isSetPPr()) {
          CTSectPr sectPr = p.getCTP().getPPr().getSectPr();
          if (sectPr != null) {
            sectionCounter++;
          }
          if (sectionCounter == 2 && sectPr != null) {
            // First section should be landscape pages
            XmlCursor cursor = sectPr.newCursor();
            if (cursor.toChild(DocxConstants.QNAME_PGSZ_ELEM)) {
              assertEquals("landscape", cursor.getAttributeText(DocxConstants.QNAME_OOXML_ORIENT_ATT));
              foundPageLayout = true;
              // Width is 14in
              assertEquals("20160", cursor.getAttributeText(DocxConstants.QNAME_OOXML_W_ATT));
              // Height is 8.5in
              assertEquals("12240", cursor.getAttributeText(DocxConstants.QNAME_OOXML_H_ATT));
            }
          }
        }

      }
    } while(iter.hasNext());
    assertTrue("Did not find expected section-level page layout", foundPageLayout);
  }

  @Test
  public void testTocGeneration() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-85-toc.swpx", "out/simplewpml-issue-85-toc.docx");
    Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
    XWPFParagraph p = iterator.next();
    assertNotNull(p);
    // Put remaining tests here.
  }

  @Test
  public void testComments() throws Exception {
    // Issue #44: comment container and commentStart/commentEnd markers.
    XWPFDocument doc = convert("simplewp/simplewpml-issue-44-comments.swpx", "out/issue-44-comments.docx");
    XWPFComment[] comments = doc.getComments();
    assertEquals("Expected 4 comments", 4, comments.length);
    XWPFComment first = comments[0];
    assertEquals("Pat Author", first.getAuthor());
    assertEquals("PA", first.getInitials());
    assertNotNull(first.getDate());
    assertEquals("Container comment body should have two paragraphs", 2, first.getParagraphs().size());
    assertEquals("Marker comment.", comments[1].getText().trim());
    assertEquals("Sam Reviewer", comments[1].getAuthor());
    String xml = doc.getDocument().getBody().xmlText();
    assertEquals("Range starts", 4, xml.split("commentRangeStart", -1).length - 1);
    assertEquals("Range ends", 4, xml.split("commentRangeEnd", -1).length - 1);
    assertEquals("Comment references", 4, xml.split("commentReference", -1).length - 1);
  }

  @Test
  public void testCommentEndWithoutStart() throws Exception {
    try {
      convert("simplewp/simplewpml-issue-44-comment-unmatched-end.swpx", "out/issue-44-comment-unmatched-end.docx");
      fail("Expected a failure for a commentEnd with no commentStart");
    } catch (DocxGenerationException e) {
      assertTrue(e.getMessage().contains("nostart"));
    }
  }

  @Test
  public void testTocEntryBookmarkRef() throws Exception {
    // Issue #151: tocentry/@bookmarkRef generates a PAGEREF to the bookmark's name.
    XWPFDocument doc = convert("simplewp/simplewpml-issue-151-toc-bookmarkref.swpx", "out/issue-151-toc-bookmarkref.docx");
    String xml = doc.getDocument().getBody().xmlText();
    assertTrue("Expected PAGEREF to the first bookmark's name", xml.contains("PAGEREF _Toc_first \\h"));
    assertTrue("Expected PAGEREF to a bookmark defined after the ToC", xml.contains("PAGEREF _Toc_second \\h"));
    assertFalse("Entry with arg-n=true should have no PAGEREF", xml.contains("PAGEREF h3"));
    assertEquals("Expected exactly two PAGEREF fields", 2, xml.split("PAGEREF", -1).length - 1);
    // Explicit entries: Word would overwrite them on update, so no dirty flag.
    assertFalse("ToC with explicit entries should not be marked dirty", xml.contains("w:dirty"));
  }

  @Test
  public void testTocWithoutEntriesIsDirty() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-151-toc-no-entries.swpx", "out/issue-151-toc-no-entries.docx");
    String xml = doc.getDocument().getBody().xmlText();
    assertTrue("Generated ToC should be marked dirty so Word updates it on open", xml.contains("w:dirty"));
  }

  @Test
  public void testDocumentPropertiesGeneration() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-140-document-properties.swpx", "out/simplewpml-issue-140-document-properties.docx");
    POIXMLProperties properties = doc.getProperties();
    assertNotNull("Expected a POIXMLProperties object", properties);
    
    // Core properties:
    
    CoreProperties coreProperties = properties.getCoreProperties();
    assertNotNull("Expected a CoreProperties object", coreProperties);

    String value = null;
    String expected = null;
    
    value = coreProperties.getCategory();
    expected = "test doc";
    assertNotNull("Expected a value for 'category' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    value = coreProperties.getContentStatus();
    expected = "draft";
    assertNotNull("Expected a value for 'contentStatus' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    Date dateValue = coreProperties.getCreated();
    Date expectedDate = isoDateFormatter.parse("2024-04-20T10:11:12Z");
    assertNotNull("Expected a value for 'created' property", dateValue);
    assertEquals("Expected \"" + expectedDate + "\", got \"" + dateValue + "\"", expectedDate, dateValue);
    
    value = coreProperties.getCreator();
    expected = "Tester McTesterson";
    assertNotNull("Expected a value for 'creator' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    value = coreProperties.getDescription();
    expected = "A test of document properties";
    assertNotNull("Expected a value for 'description' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    value = coreProperties.getIdentifier();
    expected = "issue-140-test";
    assertNotNull("Expected a value for 'identifier' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    value = coreProperties.getKeywords();
    expected = "docprops testing";
    assertNotNull("Expected a value for 'keywords' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    // NOTE: There is no getLanguage() method on CoreProperties
    
    value = coreProperties.getLastModifiedByUser();
    expected = "Tester Jr.";
    assertNotNull("Expected a value for 'lastModifiedBy' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    dateValue = coreProperties.getLastPrinted();
    expectedDate = isoDateFormatter.parse("2024-04-21T10:11:12Z");
    assertNotNull("Expected a value for 'lastPrinted' property", dateValue);
    assertEquals("Expected \"" + expectedDate + "\", got \"" + dateValue + "\"", expectedDate, dateValue);
    
    dateValue = coreProperties.getModified();
    expectedDate = isoDateFormatter.parse("2024-04-21T10:10:12Z");
    assertNotNull("Expected a value for 'modified' property", dateValue);
    assertEquals("Expected \"" + expectedDate + "\", got \"" + dateValue + "\"", expectedDate, dateValue);
    
    value = coreProperties.getRevision();
    expected = "2";
    assertNotNull("Expected a value for 'revision' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    value = coreProperties.getSubject();
    expected = "subject value";
    assertNotNull("Expected a value for 'subject' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    value = coreProperties.getTitle();
    expected = "Issue 140 Test of document properties";
    assertNotNull("Expected a value for 'title' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    value = coreProperties.getVersion();
    expected = "1";
    assertNotNull("Expected a value for 'version' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    // Test extended properties:
    
    ExtendedProperties extendedProperties = properties.getExtendedProperties();
    assertNotNull("Expected an ExtendedProperties object", extendedProperties);

    value = extendedProperties.getApplication();
    expected = "Wordinator";
    assertNotNull("Expected a value for 'Application' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    value = extendedProperties.getAppVersion();
    expected = "11.2345";
    assertNotNull("Expected a value for 'AppVersion' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    int intValue = extendedProperties.getCharacters();
    int intExpected = 345;
    assertNotNull("Expected a value for 'Characters' property", intValue);
    assertEquals("Expected \"" + intExpected + "\", got \"" + intValue + "\"", intExpected, intValue);
        
    intValue = extendedProperties.getCharactersWithSpaces();
    intExpected = 1234;
    assertNotNull("Expected a value for 'CharactersWithSpaces' property", intValue);
    assertEquals("Expected \"" + intExpected + "\", got \"" + intValue + "\"", intExpected, intValue);
        
    value = extendedProperties.getCompany();
    expected = "Planet-Sized Brains";
    assertNotNull("Expected a value for 'Company' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    intValue = extendedProperties.getHiddenSlides();
    intExpected = 0;
    assertNotNull("Expected a value for 'HiddeSlides' property", intValue);
    assertEquals("Expected \"" + intExpected + "\", got \"" + intValue + "\"", intExpected, intValue);
            
    value = extendedProperties.getHyperlinkBase();
    expected = "https://url-base-value";
    assertNotNull("Expected a value for 'HyperlinkBase' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    intValue = extendedProperties.getLines();
    intExpected = 156;
    assertNotNull("Expected a value for 'Lines' property", intValue);
    assertEquals("Expected \"" + intExpected + "\", got \"" + intValue + "\"", intExpected, intValue);
                        
    value = extendedProperties.getManager();
    expected = "Ima In Charge";
    assertNotNull("Expected a value for 'Manager' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    intValue = extendedProperties.getMMClips();
    intExpected = 1;
    assertNotNull("Expected a value for 'MMClips' property", intValue);
    assertEquals("Expected \"" + intExpected + "\", got \"" + intValue + "\"", intExpected, intValue);
            
    intValue = extendedProperties.getNotes();
    intExpected = 4;
    assertNotNull("Expected a value for 'Notes' property", intValue);
    assertEquals("Expected \"" + intExpected + "\", got \"" + intValue + "\"", intExpected, intValue);
            
    value = extendedProperties.getPresentationFormat();
    expected = "PDF";
    assertNotNull("Expected a value for 'PresentationFormat' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    value = extendedProperties.getTemplate();
    expected = "the-template-name";
    assertNotNull("Expected a value for 'Template' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    // Test custom properties:
    
    CustomProperties customProperties = properties.getCustomProperties();
    assertNotNull("Expected a CustomProperties object", customProperties);
    
    // There doesn't seem to be a method to directly query the set of defined
    // properties.
    CTProperty[] props = customProperties.getUnderlyingProperties().getPropertyArray();
    assertEquals("Expected 2 properties, got " + props.length, props.length, 2);
    
    for (int i = 0; i < props.length; i++) {
		CTProperty prop = props[i];
		assertNotNull("Expected a CTProperty object for property", prop);
    }
    
    String propName = "prop-1";
    value = props[0].getLpwstr();
    expected = "value 01";
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    propName = "prop 2";
    CTProperty prop = customProperties.getProperty(propName);
    assertNotNull("Expected a CTProperty object for property '" + propName + "'", prop);
    value = prop.getLpwstr();
    expected = "value 02";
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
  }

  public void testImageFromUrl() throws Exception {
    //    String href = "https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Google_2015_logo.svg/1200px-Google_2015_logo.svg.png";
    //    URL url = new URL(href);
    //    URLConnection conn = null;
    //    conn = url.openConnection();
    //
    //    String mimeType = conn.getContentEncoding();
    //    String mimeType = null;
    //    System.out.println("mimeType=\"" + mimeType + "\"");
    //    assertNotNull(mimeType, "Expected a MIME type");
  }

  @Test
  public void testNestedTable() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-table-nested-01.swpx", "out/table-nested-01.docx");

    // first para of text
    List<IBodyElement> contents = doc.getBodyElements();
    assertEquals(2, contents.size());

    Iterator<IBodyElement> it = contents.iterator();
    IBodyElement elem = it.next();
    assertEquals(BodyElementType.PARAGRAPH, elem.getElementType());

    XWPFParagraph p = (XWPFParagraph) elem;
    assertEquals("Nested table, just to show it works.", p.getText());

    elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());

    XWPFTable t = (XWPFTable) elem;
    assertEquals(4, t.getNumberOfRows());

    XWPFTableRow row = t.getRow(1);
    assertEquals(3, row.getTableCells().size());

    XWPFTableCell cell = row.getCell(0);
    contents = cell.getBodyElements();
    assertEquals(2, contents.size());

    it = contents.iterator();
    elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());
    t = (XWPFTable) elem;
    assertEquals(2, t.getNumberOfRows());

    elem = it.next();
    assertEquals(BodyElementType.PARAGRAPH, elem.getElementType());
  }

  @Test
  public void testNestedTableWidth() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-table-nested-02.swpx", "out/table-nested-02.docx");

    // the bug was that this used to crash (issue #114), so we only do
    // a minimal check on the output. if the conversion does not crash
    // that's already a win
    List<IBodyElement> contents = doc.getBodyElements();
    assertEquals(1, contents.size());

    Iterator<IBodyElement> it = contents.iterator();
    IBodyElement elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());
  }

  @Test
  public void testKeepNext() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-143-keepnext.swpx", "out/output-issue-143-keepnext.docx");

    List<IBodyElement> contents = doc.getBodyElements();
    int i = contents.size() - 2; // 2nd to last paragraph should have keepNext
    XWPFParagraph p = (XWPFParagraph)contents.get(i);
    
    // Check that the keepNext property is set.
    CTOnOff onOff = p.getCTPPr().getKeepNext();
    assertNotNull("Expected a value for keepNext", onOff);
    String value = (String)onOff.getVal();
    assertEquals("Expected on", value, "on");
  }

  @Test
  public void testKeepLines() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-143-keeplines.swpx", "out/output-issue-143-keeplines.docx");

    List<IBodyElement> contents = doc.getBodyElements();
    int i = contents.size() - 2; // 2nd to last paragraph should have keepLines
    XWPFParagraph p = (XWPFParagraph)contents.get(i);
    
    // Check that the keepNext property is set.
    CTOnOff onOff = p.getCTPPr().getKeepLines();
    assertNotNull("Expected a value for keepLines", onOff);
    String value = (String)onOff.getVal();
    assertEquals("Expected on", value, "on");
  }

  @Test
  public void testTableEmptyCell() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-table-empty-cell.swpx", "out/table-empty-cell.docx");

    List<IBodyElement> contents = doc.getBodyElements();
    assertEquals(1, contents.size());

    Iterator<IBodyElement> it = contents.iterator();
    IBodyElement elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());

    XWPFTable t = (XWPFTable) elem;
    assertEquals(1, t.getNumberOfRows());

    XWPFTableRow row = t.getRow(0);
    assertEquals(2, row.getTableCells().size());

    XWPFTableCell cell = row.getCell(0);
    contents = cell.getBodyElements();
    assertEquals(1, contents.size());

    cell = row.getCell(1); // the empty cell
    contents = cell.getBodyElements();
    assertEquals(1, contents.size()); // used to fail with 0
  }

  @Test
  public void testNestedTableParaBeforeTable() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-table-nested-03.swpx", "out/table-nested-01.docx");

    List<IBodyElement> contents = doc.getBodyElements();
    assertEquals(1, contents.size());

    Iterator<IBodyElement> it = contents.iterator();
    IBodyElement elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());

    XWPFTable t = (XWPFTable) elem;
    assertEquals(1, t.getNumberOfRows());

    XWPFTableRow row = t.getRow(0);
    assertEquals(1, row.getTableCells().size());

    XWPFTableCell cell = row.getCell(0);
    contents = cell.getBodyElements();
    assertEquals(3, contents.size());

    it = contents.iterator();
    elem = it.next();

    assertEquals(BodyElementType.PARAGRAPH, elem.getElementType());
    assertEquals("NOTE", ((XWPFParagraph) elem).getText());

    elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());
    t = (XWPFTable) elem;
    assertEquals(1, t.getNumberOfRows());

    elem = it.next();
    assertEquals(BodyElementType.PARAGRAPH, elem.getElementType());
  }

  @Test
  public void testMultiSectionPageProps() throws Exception {
    // verifies the solution to issues #68 and #117
    XWPFDocument doc = convert("simplewp/simplewpml-multisection-01.swpx", "out/multisection-01.docx");

    List<IBodyElement> contents = doc.getBodyElements();
    assertEquals(2, contents.size());

    Iterator<IBodyElement> it = contents.iterator();

    IBodyElement elem = it.next();
    assertEquals(BodyElementType.PARAGRAPH, elem.getElementType());
    XWPFParagraph p = (XWPFParagraph) elem;
    assertEquals("This is the first page numbered in Roman lower-case", p.getText());
    assertTrue("first para lacks section properties", p.getCTPPr().isSetSectPr());

    elem = it.next();
    assertEquals(BodyElementType.PARAGRAPH, elem.getElementType());
    p = (XWPFParagraph) elem;
    assertEquals("This is the first page numbered in decimal", p.getText());
    assertFalse("second para has section properties", p.getCTPPr().isSetSectPr());

    CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
    CTPageNumber pgNum = sectPr.getPgNumType();
    assertEquals(BigInteger.valueOf(1), pgNum.getStart());
    assertEquals(STNumberFormat.Enum.forString("decimal"), pgNum.getFmt());

    // FIXME: check header & footer (a bit tricky)

    CTPageSz pageSz = sectPr.getPgSz();
    assertEquals(STPageOrientation.Enum.forString("portrait"), pageSz.getOrient());
    assertEquals(BigInteger.valueOf(11906), pageSz.getW());
    assertEquals(BigInteger.valueOf(16838), pageSz.getH());
  }

  /**
   * Checks that each section's page number format matches the format named in the section's
   * paragraph: either format="X" (all-formats) or "expect X" (xslt-formats).
   */
  private void checkSectionPageNumberFormats(XWPFDocument doc, java.util.regex.Pattern expectedPat, int expectedCount) {
    List<IBodyElement> contents = doc.getBodyElements();
    assertEquals(expectedCount, contents.size());
    for (int i = 0; i < contents.size(); i++) {
      XWPFParagraph p = (XWPFParagraph) contents.get(i);
      java.util.regex.Matcher m = expectedPat.matcher(p.getText());
      assertTrue("No expected format in \"" + p.getText() + "\"", m.find());
      String expected = m.group(1);
      CTSectPr sectPr = (i < contents.size() - 1)
          ? p.getCTPPr().getSectPr()
          : doc.getDocument().getBody().getSectPr();
      assertNotNull("Section " + i + " has no section properties", sectPr);
      assertEquals("Section " + i + " (" + p.getText() + ")",
          STNumberFormat.Enum.forString(expected), sectPr.getPgNumType().getFmt());
    }
  }

  @Test
  public void testPageNumberAllFormats() throws Exception {
    // issue #116: every value of page-number-properties/@format other than "custom"
    XWPFDocument doc = convert("simplewp/simplewpml-issue-116-all-formats.swpx", "out/issue-116-all-formats.docx");
    checkSectionPageNumberFormats(doc, java.util.regex.Pattern.compile("format=\"([A-Za-z0-9]+)\""), 60);
  }

  @Test
  public void testPageNumberAllXsltFormats() throws Exception {
    // issue #116: supported and unsupported @xslt-format values, and @xslt-format without @format
    XWPFDocument doc = convert("simplewp/simplewpml-issue-116-xslt-formats.swpx", "out/issue-116-xslt-formats.docx");
    checkSectionPageNumberFormats(doc, java.util.regex.Pattern.compile("expect ([A-Za-z0-9]+)"), 13);
  }

  @Test
  public void testPageNumberXsltFormat() throws Exception {
    // issue #116
    XWPFDocument doc = convert("simplewp/simplewpml-issue-116-xslt-format.swpx", "out/issue-116-xslt-format.docx");
    List<IBodyElement> contents = doc.getBodyElements();
    XWPFParagraph p = (XWPFParagraph) contents.get(0);
    assertEquals(STNumberFormat.Enum.forString("lowerRoman"),
        p.getCTPPr().getSectPr().getPgNumType().getFmt());
    CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
    assertEquals(STNumberFormat.Enum.forString("upperLetter"), sectPr.getPgNumType().getFmt());
  }

  public void testTableAlign() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-table-align.swpx", "out/table-align.docx");
    List<IBodyElement> contents = doc.getBodyElements();
    assertEquals(2, contents.size());

    // first table
    Iterator<IBodyElement> it = contents.iterator();
    IBodyElement elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());

    XWPFTable t = (XWPFTable) elem;
    TableRowAlign align = t.getTableAlignment();
    assertEquals(TableRowAlign.LEFT, align);

    // second table
    elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());

    t = (XWPFTable) elem;
    align = t.getTableAlignment();
    assertEquals(TableRowAlign.CENTER, align);
  }

  @Test
  public void testNestedTableBorders() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-157-nested-tables.swpx", "out/nested-tables.docx");
    List<IBodyElement> contents = doc.getBodyElements();
    assertEquals(1, contents.size());

    // outer table
    Iterator<IBodyElement> it = contents.iterator();
    IBodyElement elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());

    XWPFTable t = (XWPFTable) elem;

    // inner table
    XWPFTableCell cell = t.getRows().get(0).getTableCells().get(0);
    elem = cell.getBodyElements().get(0);
    assertEquals(BodyElementType.TABLE, elem.getElementType());
    t = (XWPFTable) elem;

    // check the border properties
    assertEquals(XWPFTable.XWPFBorderType.NONE, t.getLeftBorderType());
    assertEquals(XWPFTable.XWPFBorderType.NONE, t.getRightBorderType());
    assertEquals(XWPFTable.XWPFBorderType.NONE, t.getTopBorderType());
    assertEquals(XWPFTable.XWPFBorderType.NONE, t.getBottomBorderType());
  }

  @Test
  public void testNumberedListsRestart() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-numbered-lists.swpx", "out/numbered-lists.docx");
    XWPFNumbering numbering = doc.getNumbering();

    XWPFParagraph a1 = findParagraph(doc, "A1");
    XWPFParagraph a2 = findParagraph(doc, "A2");
    XWPFParagraph b1 = findParagraph(doc, "B1");
    XWPFParagraph b2 = findParagraph(doc, "B2");
    assertNotNull("No numId for A1", a1.getNumID());
    assertEquals("Paragraphs in the same list should have the same numId", a1.getNumID(), a2.getNumID());
    assertEquals(b1.getNumID(), b2.getNumID());
    assertFalse("Different lists should have different numIds", a1.getNumID().equals(b1.getNumID()));

    // The list's numbering instance uses the style's abstract numbering and restarts at 1.
    XWPFStyle style = doc.getStyles().getStyle("ListNumber");
    BigInteger styleAbstractNumId = numbering.getNum(style.getCTStyle().getPPr().getNumPr().getNumId().getVal())
        .getCTNum().getAbstractNumId().getVal();
    for (XWPFParagraph para : new XWPFParagraph[] {a1, b1}) {
      CTNum ctNum = numbering.getNum(para.getNumID()).getCTNum();
      assertEquals(styleAbstractNumId, ctNum.getAbstractNumId().getVal());
      assertEquals(1, ctNum.sizeOfLvlOverrideArray());
      assertEquals(BigInteger.ZERO, ctNum.getLvlOverrideArray(0).getIlvl());
      assertEquals(BigInteger.ONE, ctNum.getLvlOverrideArray(0).getStartOverride().getVal());
    }
  }

  @Test
  public void testNumberedListWithoutStyleNumberingIsLeftAlone() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-numbered-lists.swpx", "out/numbered-lists-no-style-numbering.docx");
    XWPFNumbering numbering = doc.getNumbering();

    XWPFParagraph c1 = findParagraph(doc, "C1");
    assertNull("Paragraph in a list with no style numbering should not be given a numId", c1.getNumID());
    // Every numbering instance must point to an abstract numbering definition.
    for (XWPFNum num : numbering.getNums()) {
      assertNotNull("numId " + num.getCTNum().getNumId() + " has no abstractNumId", num.getCTNum().getAbstractNumId());
    }
  }

  @Test
  public void testTableGridFromAbsoluteColumnWidths() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-table-col-widths.swpx", "out/table-col-widths.docx");
    List<XWPFTable> tables = doc.getTables();
    assertEquals(3, tables.size());

    CTTblGrid grid = tables.get(0).getCTTbl().getTblGrid();
    assertNotNull("Expected a table grid", grid);
    assertEquals(2, grid.sizeOfGridColArray());
    assertEquals(BigInteger.valueOf(1440), grid.getGridColArray(0).getW());
    assertEquals(BigInteger.valueOf(2880), grid.getGridColArray(1).getW());
  }

  @Test
  public void testTableGridNotSetForPercentOrMissingColumnWidths() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-table-col-widths.swpx", "out/table-col-widths-no-grid.docx");
    List<XWPFTable> tables = doc.getTables();

    for (int i = 1; i <= 2; i++) {
      CTTblGrid grid = tables.get(i).getCTTbl().getTblGrid();
      assertTrue("Table " + i + " should not have grid columns",
          null == grid || 0 == grid.sizeOfGridColArray());
    }
    // Percent widths are still applied to the cells:
    XWPFTableCell cell = tables.get(1).getRow(0).getCell(0);
    CTTblWidth cellWidth = cell.getCTTc().getTcPr().getTcW();
    assertEquals(STTblWidth.PCT, cellWidth.getType());
    assertEquals(BigInteger.valueOf(1250), cellWidth.getW()); // Fiftieths of a percent
  }

  @Test
  public void testFitImagesToPageWidth() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-fit-images.swpx", "out/fit-images.docx");
    List<XWPFPicture> pictures = new java.util.ArrayList<>();
    List<CTInline> inlines = new java.util.ArrayList<>();
    for (XWPFParagraph para : doc.getParagraphs()) {
      for (XWPFRun run : para.getRuns()) {
        pictures.addAll(run.getEmbeddedPictures());
        for (CTDrawing drawing : run.getCTR().getDrawingArray()) {
          inlines.addAll(java.util.Arrays.asList(drawing.getInlineArray()));
        }
      }
    }
    assertEquals(4, pictures.size());
    assertEquals(4, inlines.size());

    // Section 1 is 8.5in wide with 1in margins, so the content width is 6.5in (9360 twips)
    long section1Width = 9360L * 635;
    assertPictureSize(pictures.get(0), inlines.get(0), section1Width, section1Width / 2); // Wide and asks to fit: shrunk, same proportions
    assertPictureSize(pictures.get(1), inlines.get(1), Units.toEMU(100), Units.toEMU(50)); // Fits already: not enlarged
    assertPictureSize(pictures.get(2), inlines.get(2), Units.toEMU(1200), Units.toEMU(600)); // Does not ask to fit: not changed

    // Section 2 is 11in wide with 1in margins, so 9in (12960 twips)
    long section2Width = 12960L * 635;
    assertPictureSize(pictures.get(3), inlines.get(3), section2Width, section2Width / 2);
  }

  @Test
  public void testContentWidth() {
    assertEquals("Word's defaults when there are no section properties", 9360, DocxGenerator.getContentWidthInTwips(null));

    CTSectPr sectPr = CTSectPr.Factory.newInstance();
    sectPr.addNewPgSz().setW(BigInteger.valueOf(11906)); // A4
    CTPageMar pageMar = sectPr.addNewPgMar();
    pageMar.setLeft(BigInteger.valueOf(1000));
    pageMar.setRight(BigInteger.valueOf(900));
    pageMar.setGutter(BigInteger.valueOf(100));
    assertEquals(11906 - 1000 - 900 - 100, DocxGenerator.getContentWidthInTwips(sectPr));
  }

  @Test
  public void testSectionColumns() throws Exception {
    // Issue #194
    XWPFDocument doc = convert("simplewp/simplewpml-issue-194-columns.swpx", "out/issue-194-columns.docx");
    List<CTSectPr> sectPrs = new java.util.ArrayList<>();
    for (XWPFParagraph para : doc.getParagraphs()) {
      if (para.getCTP().isSetPPr() && para.getCTP().getPPr().isSetSectPr()) {
        sectPrs.add(para.getCTP().getPPr().getSectPr());
      }
    }
    // The last section's properties are on the body:
    sectPrs.add(doc.getDocument().getBody().getSectPr());
    assertEquals("Expected three sections", 3, sectPrs.size());

    CTColumns cols = sectPrs.get(0).getCols();
    assertNotNull("Section 1 should have columns", cols);
    assertEquals(BigInteger.valueOf(2), cols.getNum());
    assertEquals(BigInteger.valueOf(720), cols.getSpace()); // 0.5in
    assertTrue("Section 1 should have a separator", isOn(cols.getSep()));

    assertFalse("Section 2 should not have columns", sectPrs.get(1).isSetCols());

    // The last section's columns must be carried over to the body-level section properties.
    cols = sectPrs.get(2).getCols();
    assertNotNull("Last section should have columns", cols);
    assertEquals(BigInteger.valueOf(3), cols.getNum());
    assertEquals(BigInteger.valueOf(360), cols.getSpace()); // 0.25in
    assertFalse("Last section should not have a separator", cols.isSetSep() && isOn(cols.getSep()));
  }

  private static boolean isOn(Object onOff) {
    String value = String.valueOf(onOff);
    return "true".equals(value) || "on".equals(value) || "1".equals(value);
  }

  @Test
  public void testContentWidthWithColumns() {
    CTSectPr sectPr = CTSectPr.Factory.newInstance();
    sectPr.addNewPgSz().setW(BigInteger.valueOf(12240));
    CTPageMar pageMar = sectPr.addNewPgMar();
    pageMar.setLeft(BigInteger.valueOf(1440));
    pageMar.setRight(BigInteger.valueOf(1440));
    // Content width is 9360. Two columns with the default 720 space: (9360 - 720) / 2
    CTColumns cols = sectPr.addNewCols();
    cols.setNum(BigInteger.valueOf(2));
    assertEquals(4320, DocxGenerator.getContentWidthInTwips(sectPr));
    // Three columns with 360 space: (9360 - 2 * 360) / 3
    cols.setNum(BigInteger.valueOf(3));
    cols.setSpace(BigInteger.valueOf(360));
    assertEquals(2880, DocxGenerator.getContentWidthInTwips(sectPr));
    // One column is the full width
    cols.setNum(BigInteger.valueOf(1));
    assertEquals(9360, DocxGenerator.getContentWidthInTwips(sectPr));
  }

  @Test
  public void testAnchoredImages() throws Exception {
    // Issue #195
    XWPFDocument doc = convert("simplewp/simplewpml-issue-195-anchored-images.swpx", "out/issue-195-anchored-images.docx");
    List<CTAnchor> anchors = new java.util.ArrayList<>();
    List<CTInline> inlines = new java.util.ArrayList<>();
    List<XWPFParagraph> paras = new java.util.ArrayList<>(doc.getParagraphs());
    for (XWPFTable table : doc.getTables()) {
      for (XWPFTableRow row : table.getRows()) {
        for (XWPFTableCell cell : row.getTableCells()) {
          paras.addAll(cell.getParagraphs());
        }
      }
    }
    for (XWPFParagraph para : paras) {
      for (XWPFRun run : para.getRuns()) {
        for (CTDrawing drawing : run.getCTR().getDrawingArray()) {
          anchors.addAll(java.util.Arrays.asList(drawing.getAnchorArray()));
          inlines.addAll(java.util.Arrays.asList(drawing.getInlineArray()));
        }
      }
    }
    assertEquals("Expected three body anchors and one in the table", 4, anchors.size());
    assertEquals("Expected the one image without an anchor to stay inline", 1, inlines.size());

    // Square wrap, aligned right of the margin, 0.1in (91440 EMU) from the text
    CTAnchor a = anchors.get(0);
    assertTrue(a.isSetWrapSquare());
    assertEquals(STWrapText.BOTH_SIDES, a.getWrapSquare().getWrapText());
    assertEquals(STRelFromH.MARGIN, a.getPositionH().getRelativeFrom());
    assertEquals(STAlignH.RIGHT, a.getPositionH().getAlign());
    assertEquals(STRelFromV.PARAGRAPH, a.getPositionV().getRelativeFrom());
    assertEquals(0, a.getPositionV().getPosOffset());
    assertEquals(91440L, a.getDistL());
    assertEquals(91440L, a.getDistT());
    assertFalse(a.getBehindDoc());
    assertTrue("Floating image should keep its size", a.getExtent().getCx() > 0 && a.getExtent().getCy() > 0);
    assertEquals("Same size as an inline image would be", Units.toEMU(150), a.getExtent().getCx());
    assertNotNull("Floating image should keep its picture", a.getGraphic());
    assertNotNull(a.getDocPr());

    // Top and bottom, centered
    a = anchors.get(1);
    assertTrue(a.isSetWrapTopAndBottom());
    assertEquals(STAlignH.CENTER, a.getPositionH().getAlign());
    assertEquals(91440, a.getPositionV().getPosOffset());

    // Behind text, no wrap, offsets from the margin and paragraph (1in right, 0.1in above)
    CTAnchor behind = anchors.get(2);
    assertTrue(behind.isSetWrapNone());
    assertTrue(behind.getBehindDoc());
    assertEquals(STRelFromH.MARGIN, behind.getPositionH().getRelativeFrom());
    assertEquals(914400, behind.getPositionH().getPosOffset());
    assertEquals(STRelFromV.PARAGRAPH, behind.getPositionV().getRelativeFrom());
    assertEquals(-91440, behind.getPositionV().getPosOffset());
    assertTrue("Later images should be in front of earlier ones", behind.getRelativeHeight() > anchors.get(0).getRelativeHeight());

    // In a table cell, defaults for what is not specified
    a = anchors.get(3);
    assertEquals(STRelFromH.COLUMN, a.getPositionH().getRelativeFrom());
    assertEquals(STAlignH.RIGHT, a.getPositionH().getAlign());
    assertEquals(0L, a.getDistL());
    assertEquals(STWrapText.BOTH_SIDES, a.getWrapSquare().getWrapText());
  }

  @Test
  public void testAnchoredImagesInHeaderAndFooter() throws Exception {
    // Issue #195
    XWPFDocument doc = convert("simplewp/simplewpml-issue-195-header-footer-images.swpx", "out/issue-195-header-footer-images.docx");
    assertEquals(1, doc.getHeaderList().size());
    assertEquals(1, doc.getFooterList().size());

    // Header: floating image aligned to the right margin
    XWPFHeader header = doc.getHeaderList().get(0);
    List<CTAnchor> anchors = new java.util.ArrayList<>();
    List<CTInline> inlines = new java.util.ArrayList<>();
    collectDrawings(header.getParagraphs(), anchors, inlines);
    assertEquals(1, anchors.size());
    assertEquals(0, inlines.size());
    CTAnchor a = anchors.get(0);
    assertTrue(a.isSetWrapSquare());
    assertEquals(STRelFromH.MARGIN, a.getPositionH().getRelativeFrom());
    assertEquals(STAlignH.RIGHT, a.getPositionH().getAlign());
    assertEquals("The picture data must be in the header part", 1, header.getAllPictures().size());

    // Footer: floating image centered on the page behind the text, plus an inline image
    XWPFFooter footer = doc.getFooterList().get(0);
    anchors.clear();
    inlines.clear();
    collectDrawings(footer.getParagraphs(), anchors, inlines);
    assertEquals(1, anchors.size());
    assertEquals(1, inlines.size());
    a = anchors.get(0);
    assertTrue(a.isSetWrapNone());
    assertTrue(a.getBehindDoc());
    assertEquals(STRelFromH.PAGE, a.getPositionH().getRelativeFrom());
    assertEquals(STAlignH.CENTER, a.getPositionH().getAlign());
    assertEquals(STRelFromV.PAGE, a.getPositionV().getRelativeFrom());
    assertEquals(STAlignV.CENTER, a.getPositionV().getAlign());
    // Both images use the same file, which POI stores once:
    assertEquals("The picture data must be in the footer part", 1, footer.getAllPictures().size());
  }

  private static void collectDrawings(List<XWPFParagraph> paras, List<CTAnchor> anchors, List<CTInline> inlines) {
    for (XWPFParagraph para : paras) {
      for (XWPFRun run : para.getRuns()) {
        for (CTDrawing drawing : run.getCTR().getDrawingArray()) {
          anchors.addAll(java.util.Arrays.asList(drawing.getAnchorArray()));
          inlines.addAll(java.util.Arrays.asList(drawing.getInlineArray()));
        }
      }
    }
  }

  @Test
  public void testImageWithOnlyHeightKeepsProportions() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-image-height-only.swpx", "out/image-height-only.docx");
    XWPFPicture picture = doc.getParagraphs().get(0).getRuns().get(0).getEmbeddedPictures().get(0);
    CTPositiveSize2D ext = picture.getCTPicture().getSpPr().getXfrm().getExt();
    assertEquals(Units.toEMU(50), ext.getCy());
    assertEquals("The image is square, so the width should match the height", Units.toEMU(50), ext.getCx());
  }

  @Test
  public void testTableBorderWidth() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-table-border-width.swpx", "out/table-border-width.docx");
    List<IBodyElement> contents = doc.getBodyElements();
    assertEquals(1, contents.size());

    Iterator<IBodyElement> it = contents.iterator();
    IBodyElement elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());

    XWPFTable t = (XWPFTable) elem;

    // check the border widths
    assertEquals(64, t.getLeftBorderSize());
    assertEquals(23, t.getRightBorderSize());
    assertEquals(8, t.getTopBorderSize());
    assertEquals(227, t.getBottomBorderSize());
  }

  public void testTableCellBorderWidth() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-table-cell-border-width.swpx", "out/table-cell-border-width.docx");
    List<IBodyElement> contents = doc.getBodyElements();
    assertEquals(1, contents.size());

    Iterator<IBodyElement> it = contents.iterator();
    IBodyElement elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());

    XWPFTable t = (XWPFTable) elem;
    XWPFTableCell cell = t.getRows().get(0).getTableCells().get(0);
    CTTc c = cell.getCTTc();
    CTTcPr pr = c.getTcPr();
    CTTcBorders borders = pr.getTcBorders();

    // check the border widths
    assertEquals(64, borders.getLeft().getSz().intValueExact());
    assertEquals(23, borders.getRight().getSz().intValueExact());
    assertEquals(8, borders.getTop().getSz().intValueExact());
    assertEquals(227, borders.getBottom().getSz().intValueExact());
  }

  @Test
  public void testSVG() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-svg.swpx", "out/svg.docx");
    XWPFPicture picture = getOnlyPicture(doc);
    // The SVG's root element specifies width="300" height="300"
    assertEquals("Expected width of 300", 300.0, picture.getWidth());
    assertEquals("Expected height (depth) of 300", 300.0, picture.getDepth());
  }

  @Test
  public void testSVGNoDims() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-svg-no-dims.swpx", "out/svg-no-dims.docx");
    XWPFPicture picture = getOnlyPicture(doc);
    // The SVG has no width or height, so the 200x200 default applies
    assertEquals("Expected width of 200", 200.0, picture.getWidth());
    assertEquals("Expected height (depth) of 200", 200.0, picture.getDepth());
  }

  private XWPFPicture getOnlyPicture(XWPFDocument doc) {
    List<IBodyElement> contents = doc.getBodyElements();
    assertEquals(1, contents.size());
    XWPFParagraph p = (XWPFParagraph) contents.get(0);
    assertEquals(2, p.getRuns().size());
    assertEquals("Image 1", p.getRuns().get(0).getText(0));
    List<XWPFPicture> pictures = p.getRuns().get(1).getEmbeddedPictures();
    assertEquals("Expected a picture", 1, pictures.size());
    return pictures.get(0);
  }

  // ===== INTERNAL UTILITIES

  private XWPFDocument convert(String infile, String outfile) throws Exception {
    ClassLoader classLoader = getClass().getClassLoader();
    File inFile = new File(classLoader.getResource(infile).getFile());

    File outFile = new File(outfile);
    File outDir = outFile.getParentFile();
    if (!outDir.exists()) {
      assertTrue("Failed to create directories for output file " + outFile.getAbsolutePath(), outFile.mkdirs());
    }
    if (outFile.exists()) {
      assertTrue("Failed to delete output file " + outFile.getAbsolutePath(), outFile.delete());
    }


    File templateFile = new File(classLoader.getResource(TestDocxGenerator.DOTX_TEMPLATE_PATH).getFile());
    XWPFDocument templateDoc = new XWPFDocument(new FileInputStream(templateFile));
    DocxGenerator maker = new DocxGenerator(inFile, outFile, templateDoc);
    XmlObject xml = XmlObject.Factory.parse(inFile);
    maker.generate(xml); // FIXME: why do we need to pass inFile one more time?

    FileInputStream inStream = new FileInputStream(outFile);
    XWPFDocument doc = new XWPFDocument(inStream);
    assertNotNull(doc);
    return doc;
  }

  private XWPFParagraph findParagraph(XWPFDocument doc, String text) {
    for (XWPFParagraph para : doc.getParagraphs()) {
      if (text.equals(para.getText())) {
        return para;
      }
    }
    fail("Paragraph '" + text + "' not found");
    return null;
  }

  private void assertPictureSize(XWPFPicture picture, CTInline inline, long expectedWidthEmu, long expectedHeightEmu) {
    // Both the picture's own size and the size of its inline have to match
    CTPositiveSize2D ext = picture.getCTPicture().getSpPr().getXfrm().getExt();
    assertEquals(expectedWidthEmu, ext.getCx());
    assertEquals(expectedHeightEmu, ext.getCy());
    assertEquals(expectedWidthEmu, inline.getExtent().getCx());
    assertEquals(expectedHeightEmu, inline.getExtent().getCy());
  }

}
