package org.wordinator.xml2docx;

import java.io.File;
import java.io.FileInputStream;
import java.math.BigInteger;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.wordinator.xml2docx.generator.DocxConstants;
import org.wordinator.xml2docx.generator.DocxGenerator;

import junit.framework.TestCase;

public class TestDocxGenerator extends TestCase {
	
	
	private static final String DOTX_TEMPLATE_PATH = "docx/Test_Template.dotx";

	@Test
	public void testMakeDocx() throws Exception {
		ClassLoader classLoader = getClass().getClassLoader();
		File inFile = new File(classLoader.getResource("simplewp/simplewpml-test-01.swpx").getFile());
		File templateFile = new File(classLoader.getResource(DOTX_TEMPLATE_PATH).getFile());
		File outFile = new File("out/testMakeDocx.docx");
		File outDir = outFile.getParentFile();
		System.out.println("Input file: " + inFile.getAbsolutePath());
		System.out.println("Output file: " + outFile.getAbsolutePath());
		if (!outDir.exists()) {
			assertTrue("Failed to create directories for output file " + outFile.getAbsolutePath(), outFile.mkdirs());			
		}
		if (outFile.exists()) {
			assertTrue("Failed to delete output file " + outFile.getAbsolutePath(), outFile.delete());
		}
		
		XWPFDocument templateDoc = new XWPFDocument(new FileInputStream(templateFile));
		DocxGenerator maker = new DocxGenerator(inFile, outFile, templateDoc);
		// Generate the DOCX file:
		
		try {
			XmlObject xml = XmlObject.Factory.parse(inFile);

			maker.generate(xml);
			assertTrue("DOCX file does not exist", outFile.exists());
			FileInputStream inStream = new FileInputStream(outFile);
			XWPFDocument doc = new XWPFDocument(inStream);
			assertNotNull(doc);
			Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
			XWPFParagraph p = iterator.next();
			assertNotNull("Expected a paragraph", p);
			assertEquals("Heading 1 Text", p.getText());
			System.out.println("Paragraph text='" + p.getText() + "'");
			while (iterator.hasNext()) {
			  p = iterator.next();
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
			}
			
		} catch (Exception e) {
			e.printStackTrace();
			fail("Got unexpected " + e.getClass().getSimpleName() + ": " + e.getMessage());
		}
		
	}
	
  @Test
	public void testMakeDocxWithSections() throws Exception {
		ClassLoader classLoader = getClass().getClassLoader();
		File inFile = new File(classLoader.getResource("simplewp/simplewpml-test-02.swpx").getFile());
		File templateFile = new File(classLoader.getResource(DOTX_TEMPLATE_PATH).getFile());
		File outFile = new File("out/output-02.docx");
		File outDir = outFile.getParentFile();
		System.out.println("Input file: " + inFile.getAbsolutePath());
		System.out.println("Output file: " + outFile.getAbsolutePath());
		if (!outDir.exists()) {
			assertTrue("Failed to create directories for output file " + outFile.getAbsolutePath(), outFile.mkdirs());			
		}
		if (outFile.exists()) {
			assertTrue("Failed to delete output file " + outFile.getAbsolutePath(), outFile.delete());
		}
		
		XWPFDocument templateDoc = new XWPFDocument(new FileInputStream(templateFile));
		
		DocxGenerator maker = new DocxGenerator(inFile, outFile, templateDoc);
		// Generate the DOCX file:
		
		try {
			XmlObject xml = XmlObject.Factory.parse(inFile);

			maker.generate(xml);
			assertTrue("DOCX file does not exist", outFile.exists());
			FileInputStream inStream = new FileInputStream(outFile);
			XWPFDocument doc = new XWPFDocument(inStream);
			assertNotNull(doc);
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

		} catch (Exception e) {
			e.printStackTrace();
			fail("Got unexpected " + e.getClass().getSimpleName() + ": " + e.getMessage());
		}
		
	}

  @Test
  public void testCopyNumberingDefinitions() throws Exception {
    ClassLoader classLoader = getClass().getClassLoader();
    File inFile = new File(classLoader.getResource("simplewp/simplewpml-test-03.swpx").getFile());
    File templateFile = new File(classLoader.getResource(DOTX_TEMPLATE_PATH).getFile());
    File outFile = new File("out/output-03.docx");
    File outDir = outFile.getParentFile();
    System.out.println("Input file: " + inFile.getAbsolutePath());
    System.out.println("Output file: " + outFile.getAbsolutePath());
    if (!outDir.exists()) {
      assertTrue("Failed to create directories for output file " + outFile.getAbsolutePath(), outFile.mkdirs());      
    }
    if (outFile.exists()) {
      assertTrue("Failed to delete output file " + outFile.getAbsolutePath(), outFile.delete());
    }
    
    XWPFDocument templateDoc = new XWPFDocument(new FileInputStream(templateFile));
    
    DocxGenerator maker = new DocxGenerator(inFile, outFile, templateDoc);
    // Generate the DOCX file:
    
    try {
      XmlObject xml = XmlObject.Factory.parse(inFile);

      maker.generate(xml);
      assertTrue("DOCX file does not exist", outFile.exists());
      FileInputStream inStream = new FileInputStream(outFile);
      XWPFDocument doc = new XWPFDocument(inStream);
      assertNotNull(doc);
      Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
      XWPFParagraph p = iterator.next();
      assertNotNull("Expected a paragraph", p);
      assertEquals("Test of List Formatting", p.getText());
      System.out.println("Paragraph text='" + p.getText() + "'");
      XWPFNumbering numbering = doc.getNumbering();
      assertNotNull("No numbering", numbering);

      XWPFAbstractNum abstractNumber;
      abstractNumber = numbering.getAbstractNum(BigInteger.valueOf(9));
      assertNotNull("No abstract number '9'", abstractNumber);
      
      XWPFNum num;
      num = numbering.getNum(BigInteger.valueOf(9));
      assertNotNull("No num '9'", num);
      
      
    } catch (Exception e) {
      e.printStackTrace();
      fail("Got unexpected " + e.getClass().getSimpleName() + ": " + e.getMessage());
    }
    
  }

  @Test
  public void testFootnoteGeneration() throws Exception {
    ClassLoader classLoader = getClass().getClassLoader();
    File inFile = new File(classLoader.getResource("simplewp/simplewpml-issue-29.swpx").getFile());
    File templateFile = new File(classLoader.getResource(DOTX_TEMPLATE_PATH).getFile());
    File outFile = new File("out/output-issue-29.docx");
    File outDir = outFile.getParentFile();
    System.out.println("Input file: " + inFile.getAbsolutePath());
    System.out.println("Output file: " + outFile.getAbsolutePath());
    if (!outDir.exists()) {
      assertTrue("Failed to create directories for output file " + outFile.getAbsolutePath(), outFile.mkdirs());      
    }
    if (outFile.exists()) {
      assertTrue("Failed to delete output file " + outFile.getAbsolutePath(), outFile.delete());
    }
    
    XWPFDocument templateDoc = new XWPFDocument(new FileInputStream(templateFile));
    
    DocxGenerator maker = new DocxGenerator(inFile, outFile, templateDoc);
    // Generate the DOCX file:
    
    try {
      XmlObject xml = XmlObject.Factory.parse(inFile);

      maker.generate(xml);
      assertTrue("DOCX file does not exist", outFile.exists());
      FileInputStream inStream = new FileInputStream(outFile);
      XWPFDocument doc = new XWPFDocument(inStream);
      assertNotNull(doc);
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
      
    } catch (Exception e) {
      e.printStackTrace();
      fail("Got unexpected " + e.getClass().getSimpleName() + ": " + e.getMessage());
    }
    
  }

  @Test
  public void testTableGeneration() throws Exception {
    ClassLoader classLoader = getClass().getClassLoader();
    File inFile = new File(classLoader.getResource("simplewp/simplewpml-issue-30.swpx").getFile());
    File templateFile = new File(classLoader.getResource(DOTX_TEMPLATE_PATH).getFile());
    File outFile = new File("out/output-issue-30.docx");
    File outDir = outFile.getParentFile();
    System.out.println("Input file: " + inFile.getAbsolutePath());
    System.out.println("Output file: " + outFile.getAbsolutePath());
    if (!outDir.exists()) {
      assertTrue("Failed to create directories for output file " + outFile.getAbsolutePath(), outFile.mkdirs());      
    }
    if (outFile.exists()) {
      assertTrue("Failed to delete output file " + outFile.getAbsolutePath(), outFile.delete());
    }
    
    XWPFDocument templateDoc = new XWPFDocument(new FileInputStream(templateFile));
    
    DocxGenerator maker = new DocxGenerator(inFile, outFile, templateDoc);
    try {
      XmlObject xml = XmlObject.Factory.parse(inFile);

      maker.generate(xml);
      assertTrue("DOCX file does not exist", outFile.exists());
      FileInputStream inStream = new FileInputStream(outFile);
      XWPFDocument doc = new XWPFDocument(inStream);
      assertNotNull(doc);
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
      
      assertTrue("Expected a second table", iterator.hasNext());
      table = iterator.next();

      for (XWPFTableRow row : table.getRows()) {
        // System.out.println("Row " + ++n);
        for (XWPFTableCell cell : row.getTableCells()) {
           XmlCursor cursor = cell.getCTTc().newCursor();
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
      XWPFTableRow row = table.getRow(0); // Header row.
      XWPFTableCell cell = row.getCell(1); // Center cell
      XmlCursor cursor = cell.getCTTc().newCursor();
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

    } catch (Exception e) {
      e.printStackTrace();
      fail("Got unexpected " + e.getClass().getSimpleName() + ": " + e.getMessage());
    }
  }

}
