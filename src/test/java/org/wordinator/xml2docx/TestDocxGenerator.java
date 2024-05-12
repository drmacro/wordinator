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
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
// import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.wordinator.xml2docx.generator.DocxConstants;
import org.wordinator.xml2docx.generator.DocxGenerator;

import junit.framework.TestCase;

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
  public void testDocumentPropertiesGeneration() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-issue-140-document-properties.swpx", "out/simplewpml-issue-140-document-properties.docx");
    POIXMLProperties properties = doc.getProperties();
    assertNotNull("Expected a POIXMLProperties object", properties);
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
    
    value = extendedProperties.getCompany();
    expected = "Planet-Sized Brains";
    assertNotNull("Expected a value for 'Company' property", value);
    assertEquals("Expected \"" + expected + "\", got \"" + value + "\"", expected, value);
    
    value = extendedProperties.getManager();
    expected = "Ima In Charge";
    assertNotNull("Expected a value for 'Manager' property", value);
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
    assertEquals(1, contents.size());

    it = contents.iterator();
    elem = it.next();
    assertEquals(BodyElementType.TABLE, elem.getElementType());
    t = (XWPFTable) elem;
    assertEquals(2, t.getNumberOfRows());
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

}
