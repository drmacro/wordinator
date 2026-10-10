package org.wordinator.xml2docx;

import java.io.File;
import java.nio.file.Files;
import java.util.Iterator;

import javax.xml.XMLConstants;
import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.stream.StreamSource;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;

import org.junit.Test;
import org.w3c.dom.Document;

import junit.framework.TestCase;
import net.sf.saxon.s9api.Processor;
import net.sf.saxon.s9api.QName;
import net.sf.saxon.s9api.XdmAtomicValue;
import net.sf.saxon.s9api.XdmDestination;
import net.sf.saxon.s9api.XsltTransformer;

/**
 * Tests the bundled html2docx XSLT by transforming small HTML files to SWPX
 * and checking the result with XPath.
 */
public class TestHtml2Docx extends TestCase {

  private static final String HTML2DOCX_XSL = "src/main/xsl/html2docx/html2docx.xsl";
  private static final String SWPX_NS = "urn:ns:wordinator:simplewpml";

  @Test
  public void testExternalLinksPassThrough() throws Exception {
    Document swpx = html2swpx("html/html2docx-links.html");

    assertEquals("https://example.com/a?b=c#frag",
        xpath(swpx, "string(//wp:hyperlink[wp:run = 'https']/@href)"));
    assertEquals("mailto:x@example.com",
        xpath(swpx, "string(//wp:hyperlink[wp:run = 'mail']/@href)"));
  }

  @Test
  public void testInternalLinksBecomeBookmarkReferences() throws Exception {
    Document swpx = html2swpx("html/html2docx-links.html");

    assertEquals("#top",
        xpath(swpx, "string(//wp:hyperlink[wp:run = 'fragment']/@href)"));
    // Link to another file keeps only the fragment ID.
    assertEquals("#sec1",
        xpath(swpx, "string(//wp:hyperlink[starts-with(wp:run, 'relative')]/@href)"));
  }

  @Test
  public void testOrderedAndUnorderedListStyles() throws Exception {
    Document swpx = html2swpx("html/html2docx-lists.html");

    assertEquals("List", xpath(swpx, "normalize-space(//wp:p[wp:run = 'bullet 1']/@style)"));
    assertEquals("List 2", xpath(swpx, "normalize-space(//wp:p[wp:run = 'nested bullet']/@style)"));
    assertEquals("List Number", xpath(swpx, "normalize-space(//wp:p[wp:run = 'one']/@style)"));
    assertEquals("List Number 2", xpath(swpx, "normalize-space(//wp:p[wp:run = 'nested number']/@style)"));
  }

  @Test
  public void testColumnWidthsFromColElements() throws Exception {
    Document swpx = html2swpx("html/html2docx-tables.html");

    assertEquals("25%", xpath(swpx, "string((//wp:cols)[1]/wp:col[1]/@colwidth)"));
    assertEquals("75%", xpath(swpx, "string((//wp:cols)[1]/wp:col[2]/@colwidth)"));
    // A number without units is pixels, as in HTML:
    assertEquals("100px", xpath(swpx, "string((//wp:cols)[2]/wp:col[1]/@colwidth)"));
    // A width in @style is used, ignoring other properties:
    assertEquals("2in", xpath(swpx, "string((//wp:cols)[2]/wp:col[2]/@colwidth)"));
  }

  @Test
  public void testUnusableColumnWidthsAreIgnored() throws Exception {
    Document swpx = html2swpx("html/html2docx-tables.html");

    assertEquals("4", xpath(swpx, "string(count((//wp:cols)[2]/wp:col))"));
    // HTML 4 relative width:
    assertEquals("0", xpath(swpx, "string(count((//wp:cols)[2]/wp:col[3]/@colwidth))"));
    // max-width is not width:
    assertEquals("0", xpath(swpx, "string(count((//wp:cols)[2]/wp:col[4]/@colwidth))"));
  }

  // ===== INTERNAL UTILITIES

  /**
   * Runs html2docx.xsl on a test-resource HTML file.
   * @return The generated SWPX document.
   */
  private Document html2swpx(String htmlResource) throws Exception {
    File inFile = new File(getClass().getClassLoader().getResource(htmlResource).getFile());
    File outDir = Files.createTempDirectory("html2docx-test").toFile();
    outDir.deleteOnExit();

    Processor processor = new Processor(false);
    XsltTransformer xform = processor.newXsltCompiler()
        .compile(new StreamSource(new File(HTML2DOCX_XSL)))
        .load();
    xform.setParameter(new QName("outdir"), new XdmAtomicValue(outDir.getAbsolutePath()));
    // The stylesheet writes its results with xsl:result-document, which needs a base output URI.
    xform.setBaseOutputURI(new File(outDir, "principal.out").toURI().toString());
    xform.setSource(new StreamSource(inFile));
    xform.setDestination(new XdmDestination());
    xform.transform();

    String baseName = inFile.getName().replaceAll("\\.[^.]+$", "");
    File swpxFile = new File(outDir, baseName + ".swpx");
    assertTrue("SWPX file not generated: " + swpxFile.getAbsolutePath(), swpxFile.exists());
    swpxFile.deleteOnExit();

    DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
    dbf.setNamespaceAware(true);
    return dbf.newDocumentBuilder().parse(swpxFile);
  }

  private String xpath(Document doc, String expr) throws Exception {
    XPath xp = XPathFactory.newInstance().newXPath();
    xp.setNamespaceContext(new NamespaceContext() {
      public String getNamespaceURI(String prefix) {
        return "wp".equals(prefix) ? SWPX_NS : XMLConstants.NULL_NS_URI;
      }
      public String getPrefix(String uri) { return null; }
      public Iterator<String> getPrefixes(String uri) { return null; }
    });
    return (String) xp.evaluate(expr, doc, XPathConstants.STRING);
  }

}
