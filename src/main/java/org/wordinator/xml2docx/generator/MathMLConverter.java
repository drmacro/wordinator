package org.wordinator.xml2docx.generator;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamSource;
import javax.xml.transform.Source;

import org.w3c.dom.Document;
import org.w3c.dom.Node;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import net.sf.saxon.s9api.DOMDestination;
import net.sf.saxon.s9api.Processor;
import net.sf.saxon.s9api.SaxonApiException;
import net.sf.saxon.s9api.XsltCompiler;
import net.sf.saxon.s9api.XsltExecutable;
import net.sf.saxon.s9api.XsltTransformer;

import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMathPara;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
  
public class MathMLConverter {
  private static XsltExecutable stylesheet;
  private static final Logger log = LogManager.getLogger(MathMLConverter.class.getSimpleName());

  public static void convertMath(XWPFParagraph para, XmlObject indoc) throws DocxGenerationException {
    try {
      CTOMathPara ctOMathPara = CTOMathPara.Factory.parse(convertToOOML(indoc));
    
      CTP ctp = para.getCTP();
      ctp.setOMathArray(ctOMathPara.getOMathArray());
    } catch (XmlException e) {
      // we seem to have produced bad OOXML, but this was done by the
      // stylesheet the user supplied, so treating as user error
      throw new DocxGenerationException("Could not parse CTOMathPara, bad OOXML from MathML conversion", e);
    }
  }
  
  private static Node convertToOOML(XmlObject indoc) throws DocxGenerationException {
    XsltTransformer transformer = getStylesheet().load();

    // specifying the factory, to make sure we don't get the XmlBeans
    // DOM implementation, which doesn't support level 3
    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance(
      "com.sun.org.apache.xerces.internal.jaxp.DocumentBuilderFactoryImpl",
      MathMLConverter.class.getClassLoader()
    );
    DocumentBuilder dBuilder;
    try {
      dBuilder = factory.newDocumentBuilder();
    } catch (ParserConfigurationException e) {
      // wrapping this one and letting it escape, as clearly there is some
      // serious internal problem
      throw new RuntimeException(e);
    }
    Document doc = dBuilder.newDocument();
    DOMDestination dest = new DOMDestination(doc);
    
    Source source = transformToSource(indoc);

    transformer.setSource(source);
    transformer.setDestination(dest);
    try {
      transformer.transform();
    } catch (SaxonApiException e) {
      // treating as user error, as either the user supplied a bad
      // stylesheet or the input document triggered an error
      throw new DocxGenerationException("Error converting MathML to OOXML", e);
    }

    return doc;
  }
  
  /**
   * Will load and compile the stylesheet on first call, but
   * subsequently will return the already loaded stylesheet.
   */
  private synchronized static XsltExecutable getStylesheet() throws DocxGenerationException {
    if (stylesheet != null)
      return stylesheet;

    Processor processor = new Processor(false);
    XsltCompiler compiler = processor.newXsltCompiler();
    try {
      stylesheet = compiler.compile(getTransformSource());
    } catch (SaxonApiException e) {
      // treating this as a user error, because the user is supplying
      // the stylesheet, and presumably they're the ones that messed
      // up
      throw new DocxGenerationException("XSLT compile error in MathML to OOXML stylesheet", e);
    }
    return stylesheet;
  }

  private static String XSLFILE = "MML2OMML.XSL";
  /**
   * Locate the MML2OMML.XSL file on the classpath.
   */
  private static Source getTransformSource() throws DocxGenerationException {
    InputStream is = MathMLConverter.class.getClassLoader().getResourceAsStream(XSLFILE);
    if (is == null) {
      throw new DocxGenerationException("Cannot transform MathML to OOML, because " + XSLFILE + " was not found on the classpath. See http://page.with.helpful.notes");
    }

    return new StreamSource(is);
  }

  /**
   * The XmlObject can be accessed as a DOM node, but the DOM
   * implementation doesn't support level 3, and therefore cannot be
   * processed directly by Saxon. We solve that by serializing it to
   * XML in-memory, then giving that to Saxon.
   */
  private static Source transformToSource(XmlObject indoc) {
    try {
      ByteArrayOutputStream baos = new ByteArrayOutputStream();
      indoc.save(baos);
      return new StreamSource(new ByteArrayInputStream(baos.toByteArray()));
    } catch (IOException e) {
      // if this happens, something internally is badly wrong
      throw new RuntimeException(e);
    }
  }
}
