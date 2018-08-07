package org.wordinator.xml2docx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlObject;
import org.junit.Test;
import org.wordinator.xml2docx.generator.DocxGenerator;

import junit.framework.TestCase;

public class TestDocxGenerator extends TestCase {
	
	
	private static final String DOTX_TEMPLATE_PATH = "resources/docx/Test_Template.dotx";

	@Test
	public void testMakeDocx() throws FileNotFoundException, IOException {
		ClassLoader classLoader = getClass().getClassLoader();
		File inFile = new File(classLoader.getResource("resources/simplewp/simplewpml-test-01.xml").getFile());
		File templateFile = new File(classLoader.getResource(DOTX_TEMPLATE_PATH).getFile());
		File outFile = new File("out/output.docx");
		File outDir = outFile.getParentFile();
		System.out.println("Input file: " + inFile.getAbsolutePath());
		System.out.println("Output file: " + outFile.getAbsolutePath());
		if (!outDir.exists()) {
			assertTrue("Failed to create directories for output file " + outFile.getAbsolutePath(), outFile.mkdirs());			
		}
		if (outFile.exists()) {
			assertTrue("Failed to delete output file " + outFile.getAbsolutePath(), outFile.delete());
		}
		
		DocxGenerator maker = new DocxGenerator(inFile, outFile, templateFile);
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
			
		} catch (Exception e) {
			e.printStackTrace();
			fail("Got unexpected " + e.getClass().getSimpleName() + ": " + e.getMessage());
		}
		
	}
	
  @Test
	public void testMakeDocxWithSections() throws FileNotFoundException, IOException {
		ClassLoader classLoader = getClass().getClassLoader();
		File inFile = new File(classLoader.getResource("resources/simplewp/simplewpml-test-02.xml").getFile());
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
		
		DocxGenerator maker = new DocxGenerator(inFile, outFile, templateFile);
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
				if ("First Section".equals(text)) {
					found = true;
					break;
				}
			}
			assertTrue("Did not find expected start of first section", found);
			
			// FIXME: Support for sections not yet implemented. Requires enhancing the POI API to provide for creation of
			// Paragraph-level section properties.
			// So this test should pass but should emit a warning about section-level running headers and page numbering
			// not being supported.
			
		} catch (Exception e) {
			e.printStackTrace();
			fail("Got unexpected " + e.getClass().getSimpleName() + ": " + e.getMessage());
		}
		
	}
}
