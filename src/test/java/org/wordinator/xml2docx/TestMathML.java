package org.wordinator.xml2docx;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlObject;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.wordinator.xml2docx.generator.DocxConstants;
import org.wordinator.xml2docx.generator.DocxGenerator;

import junit.framework.TestCase;

public class TestMathML extends TestCase {

  @Test
  public void testConvertMathML() throws Exception {
    XWPFDocument doc = convert("simplewp/simplewpml-mathml-01.xml", "out/testMathML.docx");

    // first para of text
    Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
    XWPFParagraph p = iterator.next();
    assertNotNull("Expected a paragraph", p);
    assertEquals("This is a line of text followed by an equation.", p.getText());

    // para with math
    p = iterator.next();
    CTP ctp = p.getCTP();
    CTOMath[] maths = ctp.getOMathArray();
    assertEquals(1, maths.length);

    // there's no point in looking into the maths to see what's there,
    // since that would amount to testing the stylesheet, but the
    // stylesheet is not part of the Wordinator
  }

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

