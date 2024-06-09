/**
 *
 */
package org.wordinator.xml2docx.generator;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.net.URLConnection;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;

import javax.imageio.ImageIO;
import javax.xml.namespace.QName;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.apache.poi.ooxml.POIXMLProperties.ExtendedProperties;
import org.apache.poi.ooxml.POIXMLProperties.CustomProperties;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFAbstractFootnoteEndnote;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFHeaderFooter;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFNum;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFSettings;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTable.XWPFBorderType;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlCursor.TokenType;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute.Space;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STOnOff1;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STVerticalAlignRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTCompat;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTCompatSetting;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFldChar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFtnEdn;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtrRef;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSettings;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STChapterSep;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STSectionMark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.wordinator.xml2docx.xwpf.model.XWPFHeaderFooterPolicy;

/**
 * Generates DOCX files from Simple Word Processing Markup Language XML.
 */
public class DocxGenerator {
  private static String NS_MATHML = "http://www.w3.org/1998/Math/MathML";

  int imageCounter = 0; // Used to keep track of count of images created.
  
  private static SimpleDateFormat isoDateFormatter = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ssX", Locale.ENGLISH);
  

  /**
   * Holds a set of table border styles
   *
   */
  protected class TableBorderStyles {

    // Default border type is set by the @borderstyle or @framestyle attribute.
    // By default there are no explicit borders.
    XWPFBorderType defaultBorderType = null;
    XWPFBorderType topBorder = null;
    XWPFBorderType bottomBorder = null;
    XWPFBorderType leftBorder = null;
    XWPFBorderType rightBorder = null;
    XWPFBorderType rowSepBorder = null;
    XWPFBorderType colSepBorder = null;

    String defaultColor = null;
    String topColor = null;
    String leftColor = null;
    String bottomColor = null;
    String rightColor = null;

    public TableBorderStyles(
        XWPFBorderType defaultBorderType,
        XWPFBorderType topBorder,
        XWPFBorderType bottomBorder,
        XWPFBorderType leftBorder,
        XWPFBorderType rightBorder) {

    }

    /**
     * Construct using specified border styles as the initial values.
     * @param parentBorderStyles Styles to be inherited from parent
     */
    public TableBorderStyles(TableBorderStyles parentBorderStyles) {
      defaultBorderType = parentBorderStyles.getDefaultBorderType();
      topBorder = parentBorderStyles.getTopBorder();
      bottomBorder = parentBorderStyles.getBottomBorder();
      leftBorder = parentBorderStyles.getLeftBorder();
      rightBorder = parentBorderStyles.getRightBorder();
      rowSepBorder = parentBorderStyles.getRowSepBorder();
      colSepBorder = parentBorderStyles.getColSepBorder();

      // Get default border colors from parent?
    }

    /**
     * Construct initial border styles from an element that may specify
     * border frame style attributes.
     * @param borderStyleSpecifier XML element that may specify frame style attributes (table, td)
     */
    public TableBorderStyles(XmlObject borderStyleSpecifier) {

      XmlCursor cursor = borderStyleSpecifier.newCursor();
      String tagname = cursor.getName().getLocalPart();
      String styleValue = null;
      String styleBottomValue= null;
      String styleTopValue= null;
      String styleLeftValue= null;
      String styleRightValue= null;

      String colorValue = null;
      String colorBottomValue= null;
      String colorTopValue= null;
      String colorLeftValue= null;
      String colorRightValue= null;

      // Issue 30: Also get the border color values.

      if ("table".equals(tagname)) {
        styleValue = cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_ATT);
        styleBottomValue= cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_BOTTOM_ATT);
        styleTopValue= cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_TOP_ATT);
        styleLeftValue= cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_LEFT_ATT);
        styleRightValue= cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_RIGHT_ATT);
      } else {
        styleValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_ATT);
        styleBottomValue= cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_BOTTOM_ATT);
        styleTopValue= cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_TOP_ATT);
        styleLeftValue= cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_LEFT_ATT);
        styleRightValue= cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_RIGHT_ATT);

        colorValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_COLOR_ATT);
        colorBottomValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_COLOR_BOTTOM_ATT);
        colorTopValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_COLOR_TOP_ATT);
        colorLeftValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_COLOR_LEFT_ATT);
        colorRightValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_COLOR_RIGHT_ATT);
      }

      if (styleValue != null) {
        setDefaultBorderType(xwpfBorderType(styleValue));
      }

      if (styleBottomValue != null) {
        setBottomBorder(xwpfBorderType(styleBottomValue));
      }
      if (styleTopValue != null) {
        setTopBorder(xwpfBorderType(styleTopValue));
      }
      if (styleLeftValue != null) {
        setLeftBorder(xwpfBorderType(styleLeftValue));
      }
      if (styleRightValue != null) {
        setRightBorder(xwpfBorderType(styleRightValue));
      }

      if (colorValue != null) {
        setDefaultBorderColor(colorValue);
      }

      if (colorBottomValue != null) {
        setBottomColor(colorBottomValue);
      }
      if (colorTopValue != null) {
        setTopColor(colorTopValue);
      }
      if (colorLeftValue != null) {
        setLeftColor(colorLeftValue);
      }
      if (colorRightValue != null) {
        setRightColor(colorRightValue);
      }
    }

    public void setDefaultBorderColor(String colorValue) {
      this.defaultColor = colorValue;
      if (this.getBottomColor() == null) this.setBottomColor(colorValue);
      if (this.getTopColor() == null) this.setTopColor(colorValue);
      if (this.getLeftColor() == null) this.setLeftColor(colorValue);
      if (this.getRightColor() == null) this.setRightColor(colorValue);

    }

    public String getBottomColor() {
      return this.bottomColor;
    }

    public String getTopColor() {
      return this.topColor;
    }

    public String getLeftColor() {
      return this.leftColor;
    }

    public String getRightColor() {
      return this.rightColor;
    }

    public void setBottomColor(String colorValue) {
      this.bottomColor = colorValue;
    }

    public void setTopColor(String colorValue) {
      this.topColor = colorValue;
    }

    public void setLeftColor(String colorValue) {
      this.leftColor = colorValue;
    }

    public void setRightColor(String colorValue) {
      this.rightColor = colorValue;
    }

    public XWPFBorderType getDefaultBorderType() {
      return defaultBorderType;
    }

    public void setDefaultBorderType(XWPFBorderType defaultBorderType) {
      this.defaultBorderType = defaultBorderType;
      if (getBottomBorder() == null) setBottomBorder(defaultBorderType);
      if (getTopBorder() == null) setTopBorder(defaultBorderType);
      if (getLeftBorder() == null) setLeftBorder(defaultBorderType);
      if (getRightBorder() == null) setRightBorder(defaultBorderType);
    }

    public XWPFBorderType getTopBorder() {
      return topBorder;
    }

    public void setTopBorder(XWPFBorderType topBorder) {
      this.topBorder = topBorder;
    }

    public XWPFBorderType getBottomBorder() {
      return bottomBorder;
    }

    public STBorder.Enum getBottomBorderEnum() {
      return getBorderEnumForType(getBottomBorder());
    }
    public STBorder.Enum getTopBorderEnum() {
      return getBorderEnumForType(getTopBorder());
    }
    public STBorder.Enum getLeftBorderEnum() {
      return getBorderEnumForType(getLeftBorder());
    }
    public STBorder.Enum getRightBorderEnum() {
      return getBorderEnumForType(getRightBorder());
    }

    public STBorder.Enum getBorderEnumForType(XWPFBorderType type) {
      STBorder.Enum result = null;
      if (type != null) {
        result = stBorderType(type);
      }
      return result;
    }

    public void setBottomBorder(XWPFBorderType bottomBorder) {
      this.bottomBorder = bottomBorder;
    }

    public XWPFBorderType getLeftBorder() {
      return leftBorder;
    }

    public void setLeftBorder(XWPFBorderType leftBorder) {
      this.leftBorder = leftBorder;
    }

    public XWPFBorderType getRightBorder() {
      return rightBorder;
    }

    public void setRightBorder(XWPFBorderType rightBorder) {
      this.rightBorder = rightBorder;
    }

    public XWPFBorderType getRowSepBorder() {
      return rowSepBorder;
    }

    public void setRowSepBorder(XWPFBorderType rowSepBorder) {
      this.rowSepBorder = rowSepBorder;
    }

    public XWPFBorderType getColSepBorder() {
      return colSepBorder;
    }

    public void setColSepBorder(XWPFBorderType colSepBorder) {
      this.colSepBorder = colSepBorder;
    }

    /**
     * Determine if any borders are explicitly set
     * @return True if one or more borders have a defined style.
     */
    public boolean hasBorders() {
      boolean result =
              getDefaultBorderType() != null ||
              getBottomBorder() != null ||
              getTopBorder() != null ||
              getLeftBorder() != null ||
              getRightBorder() != null;
      return result;
    }

  }

  private static final Logger log = LogManager.getLogger(DocxGenerator.class);

  private File outFile;
  private int dotsPerInch = 72; /* DPI */
  // Map of source IDs to internal object IDs.
  private Map<String, BigInteger> bookmarkIdToIdMap = new HashMap<String, BigInteger>();
  private int idCtr = 0;
  private File inFile;
  private XWPFDocument templateDoc;

  // Set to false when a style warning is issued.
  private boolean isFirstParagraphStyleWarning = true;
  private boolean isFirstCharacterStyleWarning = true;
  private boolean isFirstTableStyleWarning = true;

  /**
   *
   * @param inFile File representing input document.
   * @param outFile File to write DOCX result to
   * @param templateDoc DOTX template to initialize result DOCX with (provides style definitions)
   * @throws Exception Exception from loading the template document
   * @throws FileNotFoundException If the template odcument is not found
   */
  public DocxGenerator(File inFile, File outFile, XWPFDocument templateDoc) throws FileNotFoundException, Exception {
    this.inFile = inFile;
    this.outFile = outFile;
    this.templateDoc = templateDoc;
  }

  /*
   * Generate the DOCX file from the input Simple WP ML document.
   * @param xml The XmlObject that holds the Simple WP XML content
   */
  public void generate(XmlObject xml) throws DocxGenerationException, XmlException, IOException {

    XWPFDocument doc = new XWPFDocument();

    setupNumbering(doc, this.templateDoc);
    setupStyles(doc, this.templateDoc);
    constructDoc(doc, xml);

    FileOutputStream out = new FileOutputStream(outFile);
        doc.write(out);
    doc.close();
  }

  /**
   * Walk the XML document to create the Word document.
   * @param doc Word document to write to
   * @param xml Simple ML doc to walk
   */
  private void constructDoc(XWPFDocument doc, XmlObject xml) throws DocxGenerationException {
    XmlCursor cursor = xml.newCursor();
    cursor.toFirstChild(); // Put us on the root element of the document
    cursor.push();
    if (cursor.toChild(DocxConstants.SIMPLE_WP_NS, "document-properties")) {
    	handleDocumentProperties(doc, cursor.getObject());
    }
    cursor.pop();
    cursor.push();
    cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "body"));
    setDocSettings(doc, xml);
    handleBody(doc, cursor.getObject());

    cursor.pop();

    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-sequence-properties"))) {
      // Results in a w:sectPr as  the last child
      // of w:body.
      setupPageSequence(doc, cursor.getObject());
    } else {
      CTDocument1 document = doc.getDocument();
      CTBody body = (document.isSetBody() ? document.getBody() : document.addNewBody());
      @SuppressWarnings("unused")
      CTSectPr sectPr = (body.isSetSectPr() ? body.getSectPr() : body.addNewSectPr());
      // At this point let Word fill in the details.

    }
    cursor.pop();

  }

  /**
   * Set the document-level settings.
   * @param doc Word document to write to
   * @param xml Simple ML doc
   */
  private void setDocSettings(XWPFDocument doc, XmlObject xml) {
	  // Issue #133: Turn off compatibility mode.
      XWPFSettings settings = doc.getSettings();
      CTSettings ctSettings = settings.getCTSettings();
      CTCompat compat = ctSettings.addNewCompat();
      // This may be all we need to do.
      CTCompatSetting compatSetting = compat.addNewCompatSetting();
      // Name, URI, and value come from inspecting working Word docs.
      // I do not know where these values are documented.
      compatSetting.setName("compatibilityMode");
      compatSetting.setUri("http://schemas.microsoft.com/office/word");
      compatSetting.setVal("15");
  }

/**
   * Set document properties.
   * @param doc Document to add properties to.
   * @param xml <document-properties> element
   */
  private void handleDocumentProperties(
		  XWPFDocument doc, 
		  XmlObject xml) {
	  XmlCursor cursor = xml.newCursor();
	  cursor.push();
	  if (cursor.toChild(DocxConstants.QNAME_CORE_PROPERTIES_ELEM)) {
		  handleCoreProperties(doc, cursor.getObject());
	  }
	  cursor.pop();
	  cursor.push();
	  if (cursor.toChild(DocxConstants.QNAME_EXTENDED_PROPERTIES_ELEM)) {
		  handleExtendedProperties(doc, cursor.getObject());
	  }
	  cursor.pop();
	  if (cursor.toChild(DocxConstants.QNAME_CUSTOM_PROPERTIES_ELEM)) {
		  handleCustomProperties(doc, cursor.getObject());
	  }
	  
  }

/**
 * Set core properties from the <core-properties> element.
 * @param doc XWPF document to set the properties on
 * @param xml <core-properties> element.
 */
private void handleCoreProperties(XWPFDocument doc, XmlObject xml) {
	// DateTime properties have ISO times like:
	// 2022-12-18T18:05:00Z
	POIXMLProperties properties = doc.getProperties();
	CoreProperties coreProperties = properties.getCoreProperties();
	XmlCursor cursor = xml.newCursor();	
	if (cursor.toFirstChild()) {
		do {
			String tagName = cursor.getName().getLocalPart();
			String value = cursor.getTextValue();
			if ("category".equals(tagName)) {
				coreProperties.setCategory(value);
			} else if ("contentStatus".equals(tagName)) {
				coreProperties.setContentStatus(value);
			} else if ("created".equals(tagName)) {
				try {					
					Date date = isoDateFormatter.parse(value);
					Optional<Date> opional = Optional.of(date);
					coreProperties.setCreated(opional);
				} catch (Exception e) {
					log.warn("handleCoreProperties(): " + e.getClass().getSimpleName() + " parsing <created> value '" + value + "'");
				}
			} else if ("creator".equals(tagName)) {
				coreProperties.setCreator(value);
			} else if ("description".equals(tagName)) {
				coreProperties.setDescription(value);
			} else if ("identifier".equals(tagName)) {
				coreProperties.setIdentifier(value);
			} else if ("keywords".equals(tagName)) {
				coreProperties.setKeywords(value);
			} else if ("language".equals(tagName)) {
				// There doesn't see to be a setLanguage() method on CoreProperties
			} else if ("lastModifiedBy".equals(tagName)) {
				coreProperties.setLastModifiedByUser(value);
			} else if ("lastPrinted".equals(tagName)) {
				try {					
					Date date = isoDateFormatter.parse(value);
					Optional<Date> opional = Optional.of(date);
					coreProperties.setLastPrinted(opional);
				} catch (Exception e) {
					log.warn("handleCoreProperties(): " + e.getClass().getSimpleName() + " parsing <lastPrinted> value '" + value + "'");
				}				
			} else if ("modified".equals(tagName)) {
				try {					
					Date date = isoDateFormatter.parse(value);
					Optional<Date> opional = Optional.of(date);
					coreProperties.setModified(opional);
				} catch (Exception e) {
					log.warn("handleCoreProperties(): " + e.getClass().getSimpleName() + " parsing <modified> value '" + value + "'");
				}				
			} else if ("revision".equals(tagName)) {
				coreProperties.setRevision(value);
			} else if ("subject".equals(tagName)) {
				coreProperties.setSubjectProperty(value);
			} else if ("title".equals(tagName)) {
				coreProperties.setTitle(value);
			} else if ("version".equals(tagName)) {
				coreProperties.setVersion(value);
			} else {
				log.warn("handleCoreProperties(): Unexpected element '" + tagName + "' in <core-properties>. Ignored.");
			}
		} while (cursor.toNextSibling());
	}	
}

/**
 * Handle the <extended-properties> element.
 * @param doc XWPF document to set the properties on
 * @param xml <extended-properties> element.
 */
private void handleExtendedProperties(XWPFDocument doc, XmlObject xml) {
	POIXMLProperties properties = doc.getProperties();
	ExtendedProperties extendedProperties = properties.getExtendedProperties();
	XmlCursor cursor = xml.newCursor();	
	
	if (cursor.toFirstChild()) {
		do {
			String tagName = cursor.getName().getLocalPart();
			String value = cursor.getTextValue();
			if ("Application".equals(tagName)) {
				extendedProperties.setApplication(value);
			} else if ("AppVersion".equals(tagName)) {
				extendedProperties.setAppVersion(value);
			} else if ("Characters".equals(tagName)) {
				try {
					int intValue = Integer.parseInt(value);
					extendedProperties.setCharacters(intValue);
				} catch (Exception e) {
					log.warn("handleExtendedProperties(): " + e.getClass().getSimpleName() + " parsing <" + tagName + "> value '" + value + "'");
				}
			} else if ("CharactersWithSpaces".equals(tagName)) {
				try {
					int intValue = Integer.parseInt(value);
					extendedProperties.setCharactersWithSpaces(intValue);
				} catch (Exception e) {
					log.warn("handleExtendedProperties(): " + e.getClass().getSimpleName() + " parsing <" + tagName + "> value '" + value + "'");
				}
			} else if ("Company".equals(tagName)) {
				extendedProperties.setCompany(value);
			} else if ("DigSig".equals(tagName)) {
				log.warn("handleExtendedProperties(): No set method for '" + tagName + "' extended property.");
			} else if ("DocSecurity".equals(tagName)) {
				log.warn("handleExtendedProperties(): No set method for '" + tagName + "' extended property.");
			} else if ("HeadingPairs".equals(tagName)) {
				log.warn("handleExtendedProperties(): No set method for '" + tagName + "' extended property.");
			} else if ("HiddenSlides".equals(tagName)) {
				try {
					int intValue = Integer.parseInt(value);
					extendedProperties.setHiddenSlides(intValue);
				} catch (Exception e) {
					log.warn("handleExtendedProperties(): " + e.getClass().getSimpleName() + " parsing <" + tagName + "> value '" + value + "'");
				}
			} else if ("HLinks".equals(tagName)) {
				log.warn("handleExtendedProperties(): No set method for '" + tagName + "' extended property.");
			} else if ("HyperlinkBase".equals(tagName)) {
				extendedProperties.setHyperlinkBase(value);
			} else if ("HyperlinksChanged".equals(tagName)) {
				log.warn("handleExtendedProperties(): No set method for '" + tagName + "' extended property.");
			} else if ("Lines".equals(tagName)) {
				try {
					int intValue = Integer.parseInt(value);
					extendedProperties.setLines(intValue);
				} catch (Exception e) {
					log.warn("handleExtendedProperties(): " + e.getClass().getSimpleName() + " parsing <" + tagName + "> value '" + value + "'");
				}
			} else if ("LinksUpToDate".equals(tagName)) {
				log.warn("handleExtendedProperties(): No set method for '" + tagName + "' extended property.");
			} else if ("Manager".equals(tagName)) {
				extendedProperties.setManager(value);
			} else if ("MMClips".equals(tagName)) {
				try {
					int intValue = Integer.parseInt(value);
					extendedProperties.setMMClips(intValue);
				} catch (Exception e) {
					log.warn("handleExtendedProperties(): " + e.getClass().getSimpleName() + " parsing <" + tagName + "> value '" + value + "'");
				}
			} else if ("Notes".equals(tagName)) {
				try {
					int intValue = Integer.parseInt(value);
					extendedProperties.setNotes(intValue);
				} catch (Exception e) {
					log.warn("handleExtendedProperties(): " + e.getClass().getSimpleName() + " parsing <" + tagName + "> value '" + value + "'");
				}
			} else if ("Paragraphs".equals(tagName)) {
				try {
					int intValue = Integer.parseInt(value);
					extendedProperties.setParagraphs(intValue);
				} catch (Exception e) {
					log.warn("handleExtendedProperties(): " + e.getClass().getSimpleName() + " parsing <" + tagName + "> value '" + value + "'");
				}
			} else if ("PresentationFormat".equals(tagName)) {
				extendedProperties.setPresentationFormat(value);
			} else if ("ScaleCrop".equals(tagName)) {
				log.warn("handleExtendedProperties(): No set method for '" + tagName + "' extended property.");
			} else if ("SharedDoc".equals(tagName)) {
				log.warn("handleExtendedProperties(): No set method for '" + tagName + "' extended property.");
			} else if ("Slides".equals(tagName)) {
				try {
					int intValue = Integer.parseInt(value);
					extendedProperties.setParagraphs(intValue);
				} catch (Exception e) {
					log.warn("handleExtendedProperties(): " + e.getClass().getSimpleName() + " parsing <" + tagName + "> value '" + value + "'");
				}
			} else if ("Template".equals(tagName)) {
				extendedProperties.setTemplate(value);
			} else if ("TitlesOfParts".equals(tagName)) {
				log.warn("handleExtendedProperties(): No set method for '" + tagName + "' extended property.");
			} else if ("TotalTime".equals(tagName)) {
				try {
					int intValue = Integer.parseInt(value);
					extendedProperties.setTotalTime(intValue);
				} catch (Exception e) {
					log.warn("handleExtendedProperties(): " + e.getClass().getSimpleName() + " parsing <" + tagName + "> value '" + value + "'");
				}
			} else if ("contentStatus".equals(tagName)) {
				extendedProperties.setApplication(value);
			} else if ("contentStatus".equals(tagName)) {
				extendedProperties.setApplication(value);
			} else if ("Words".equals(tagName)) {
				try {
					int intValue = Integer.parseInt(value);
					extendedProperties.setWords(intValue);
				} catch (Exception e) {
					log.warn("handleExtendedProperties(): " + e.getClass().getSimpleName() + " parsing <" + tagName + "> value '" + value + "'");
				}
			}
		} while (cursor.toNextSibling());
	}	
}

private void handleCustomProperties(XWPFDocument doc, XmlObject xml) {
	POIXMLProperties properties = doc.getProperties();
	CustomProperties customProperties = properties.getCustomProperties();
	
	XmlCursor cursor = xml.newCursor();	
	if (cursor.toFirstChild()) {
		do {
			String tagName = cursor.getName().getLocalPart();
			String value = cursor.getTextValue();
			String propName = cursor.getAttributeText(DocxConstants.QNAME_NAME_ATT);
			if (propName == null || "".equals(propName)) {
				log.warn("handleCustomProperties(): No value for required @name attribute on <" + tagName + "> element with value '" + value + "'");
				continue;				
			}
			customProperties.addProperty(propName, value);
		} while (cursor.toNextSibling());
	}
}


/**
   * Process the elements in &lt;body&gt;
   * @param doc Document to add paragraphs to.
   * @param xml Body or section element
   * @return Last paragraph of the body (if any)
   * @throws DocxGenerationException
   */
  private XWPFParagraph handleBody(
      XWPFDocument doc,
      XmlObject xml)
          throws DocxGenerationException {
    if (log.isDebugEnabled()) {
      // log.debug("handleBody(): starting...");
    }
    XWPFParagraph lastPara = null;
    XmlCursor cursor = xml.newCursor();
    if (cursor.toFirstChild()) {
      do {
        lastPara = null;
        String tagName = cursor.getName().getLocalPart();
        String namespace = cursor.getName().getNamespaceURI();
        if ("p".equals(tagName)) {
          XWPFParagraph p = doc.createParagraph();
          makeParagraph(p, cursor);
          lastPara = p;
        } else if ("section".equals(tagName)) {
          handleSection(doc, cursor.getObject());
        } else if ("table".equals(tagName)) {
          XWPFTable table = doc.createTable();
          makeTable(table, cursor.getObject());
        } else if ("object".equals(tagName)) {
          // FIXME: This is currently unimplemented.
          makeObject(doc, cursor);
        } else if ("toc".equals(tagName)) {
          makeTableOfContents(doc, cursor.getObject());
        } else {
          log.warn("handleBody(): Unexpected element {" + namespace + "}:'" + tagName + "' in <body>. Ignored.");
        }
      } while (cursor.toNextSibling());

    }
    return lastPara;
  }

  /**
   * Count the number of section elements within the document body
   * Issue 51: Created this then decided it wasn't needed. Keeping the method in case
   * it's useful in the future.
   * @param cursor A cursor created from the body element to count the sections in.
   * @return The number of sections found.
   */
  @SuppressWarnings("unused")
  private int countSections(XmlCursor cursor) {
    int count = 0;
    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "section"))) {
      count++;
      while (cursor.toNextSibling(new QName(DocxConstants.SIMPLE_WP_NS, "section"))) {
        count++;
      }
    }
    return count;
  }

  /**
   * Generate a table of contents field.
   * @param doc Document we're adding to
   * @param xml &lt;toc&gt; element
   */
  private void makeTableOfContents(
      XWPFDocument doc,
      XmlObject xml)
          throws DocxGenerationException
  {
    XmlCursor cursor = xml.newCursor();
    cursor.push();
    XWPFParagraph para = doc.createParagraph();
    String tocStyleId = "TOC1";
    XWPFStyle tocStyle = para.getDocument().getStyles().getStyle(tocStyleId);
    if (tocStyle == null) {
      makeParagraphStyle(doc, tocStyleId, tocStyleId);
    }
    para.setStyle(tocStyleId);

    // Start the outer field

    makeTocStartField(cursor, para);

    para.createRun().getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);

    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "tocentry"))) {
      int tocLevel = 1;
      do {
    	  handleTocEntry(doc, cursor, tocLevel);
      } while (cursor.toNextSibling());
    }
    cursor.pop();
    para = doc.createParagraph();
    para.createRun().getCTR().addNewFldChar().setFldCharType(STFldCharType.END);

  }

  /**
   * Handles a table of contents entry
   * @param doc Document to add the ToC entry to
   * @param cursor Cursor pointing at <tocentry> element
   * @param tocLevel The current ToC level. 1 (one) = highest level
   * @throws Exception
   */
  private void handleTocEntry(XWPFDocument doc, XmlCursor cursor, int tocLevel) throws DocxGenerationException {

    // The tocentry element can contain a <p> that provides the text of the toc entry
    // It can also contain nested <tocentry> elements.

    // With POI 5 and XML Beans 4 we can use XPath and XQuery with Saxon 10 but with
    // POI 4 we would need Saxon 9.0.0.4 and don't want to include that in the dependencies
    // since it would interfere with Saxon 10 as far as I know.
    // So for now just using a simple walk of the children.

    cursor.push();


    XWPFParagraph para = doc.createParagraph();
    cursor.push();
    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "p"))) {
      this.makeParagraph(para, cursor);
    }
    cursor.pop();

    // Set to TOC style
    String tocStyleId = "TOC" + tocLevel;
    XWPFStyle tocStyle = para.getDocument().getStyles().getStyle(tocStyleId);
    if (tocStyle == null) {
      makeParagraphStyle(doc, tocStyleId, tocStyleId);
    }
    para.setStyle(tocStyleId);

    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "tocentry"))) {
      cursor.push();
      do {
        // Assuming all children are tocentry from here on.
        handleTocEntry(doc, cursor, tocLevel + 1);
      } while (cursor.toNextSibling());
      cursor.pop();
    }

    cursor.pop();

  }

  /**
   * Makes the start field for a table of contents complex field.
   *
   * @param cursor The XML cursor pointing at a <toc> element.   *
   * @param para The paragraph that will contain the field.
   */
  private void makeTocStartField(XmlCursor cursor, XWPFParagraph para) {
    // Issue 85: Set w:dirty="true" on the field to trigger ToC update on open
    //           when the Word Update automatic links on open setting is active.
    CTFldChar field = para.createRun().getCTR().addNewFldChar();
    field.setFldCharType(STFldCharType.BEGIN);
    field.setDirty(STOnOff1.ON);
    CTText ctText = para.createRun().getCTR().addNewInstrText();
    ctText.setSpace(Space.PRESERVE);
    String tocOptions = "";
    String attValue = null;

    // Includes captioned items, but omits caption labels and numbers.
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_A_ATT);
    if (null != attValue) {
      tocOptions += " \\a \"" + attValue + "\"";
    } // No default

    // Includes entries only from the portion of the document marked by
    // the named bookmark
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_B_ATT);
    if (null != attValue) {
      tocOptions += " \\b \"" + attValue + "\"";
    } // No default

    // Includes figures, tables, charts, and other items that are numbered
    // by a SEQ field
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_C_ATT);
    if (null != attValue) {
      tocOptions += " \\c \"" + attValue + "\"";
    } // No default

    // Sequence separator character
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_D_ATT);
    if (null != attValue) {
      tocOptions += " \\d \"" + attValue + "\"";
    } // No default

    // The list type to include
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_F_ATT);
    if (null != attValue) {
      tocOptions += " \\f \"" + attValue + "\"";
    }

    // Use hyperlinks
    // Default is "true" per the SWPX grammar
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_H_ATT);
    if ("false".equalsIgnoreCase(attValue)) {
      // Omit \h option
    } else {
      tocOptions += " \\h";
    }

    // Levels to include
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_L_ATT);
    if (null != attValue) {     
      tocOptions += " \\l \"" + attValue + "\"";
    }

    // Turn off page numbers
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_N_ATT);
    if (null != attValue) {     
      tocOptions += " \\n";
      if ("none".equalsIgnoreCase(attValue)) {
        // No additional parameter
      } else {
        // Value is a range: 2-4
        tocOptions += " \"" + attValue + "\"";
      }
    }

    // Heading levels to include
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_O_ATT);
    if (null != attValue) {
      if ("all".equalsIgnoreCase(attValue)) {
        // No additional parameter
        tocOptions += " \\o";
      } else {
        tocOptions += " \\o \"" + attValue + "\"";
      }
    }

    // Entry and page number separator
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_P_ATT);
    if (null != attValue) {
      tocOptions += " \\p \"" + attValue + "\"";
    }

    // For entries numbered with a SEQ field (§17.16.5.56), adds a prefix
    // to the page number.
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_S_ATT);
    if (null != attValue) {
      tocOptions += " \\s \"" + attValue + "\"";
    }

    // Style names to select
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_T_ATT);
    if (null != attValue) {
      tocOptions += " \\t \"" + attValue + "\"";
    }

    // Uses the applied paragraph outline level.
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_U_ATT);
    if (null != attValue) {
      if ("false".equalsIgnoreCase(attValue)) {
        // Omit \\u option
      } else {
        tocOptions += " \\u";
      }
    }
    
    // Preserves tab entries within table entries.
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_W_ATT);
    if (null != attValue) {
      if ("false".equalsIgnoreCase(attValue)) {
        // Omit \w option
      } else {
        tocOptions += " \\w";
      }
    }
    
    // Preserves newline characters within table entries.
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_X_ATT);
    if (null != attValue) {
      if ("false".equalsIgnoreCase(attValue)) {
        // Omit \x option
      } else {
        tocOptions += " \\x";
      }
    }
    
    // Hides tab leader and page numbers in web page view (§17.18.102).
    // Default is "true" per the SWPX grammar
    attValue = cursor.getAttributeText(DocxConstants.QNAME_ARG_Z_ATT);
    if ("false".equalsIgnoreCase(attValue)) {
      // Omit \z option
    } else {
      tocOptions += " \\z";
    }

    ctText.setStringValue("TOC " + tocOptions);
    
  }

  private void makeParagraphStyle(XWPFDocument doc, String styleId, String string) {
    CTStyle style = CTStyle.Factory.newInstance();
    style.setStyleId(styleId);
    style.setType(STStyleType.PARAGRAPH);
    style.addNewName().setVal(styleId);
    style.addNewBasedOn().setVal("DefaultParagraphFont");
    style.addNewUiPriority().setVal(new BigInteger("99"));
    style.addNewSemiHidden();
    style.addNewUnhideWhenUsed();
    doc.getStyles().addStyle(new XWPFStyle(style));
  }

  /**
   * Handle a &lt;section&gt; element
   * @param doc Document we're adding to
   * @param xml &lt;section&gt; element
   */
  private void handleSection(
      XWPFDocument doc,
      XmlObject xml)
          throws DocxGenerationException {
    XmlCursor cursor = xml.newCursor();

    XmlObject localPageSequenceProperties = null;

    cursor.push();
    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-sequence-properties"))) {
      localPageSequenceProperties = cursor.getObject();
    }
    cursor.pop();

    String sectionType = cursor.getAttributeText(DocxConstants.QNAME_TYPE_ATT);

    cursor.push();
    cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "body"));

    XWPFParagraph lastPara = handleSectionContent(doc, cursor.getObject());

    CTSectPr sectPr = null;

//    if (log.isDebugEnabled()) {
//      log.debug("handleSection(): Setting sectPr on last paragraph.");
//    }
    CTPPr ppr = (lastPara.getCTP().isSetPPr() ? lastPara.getCTP().getPPr() : lastPara.getCTP().addNewPPr());
    sectPr = ppr.addNewSectPr();
    ppr.setSectPr(sectPr);

    if (sectionType != null) {
      CTSectType type = sectPr.addNewType();
      type.setVal(STSectionMark.Enum.forString(sectionType));
    }

    if (localPageSequenceProperties != null) {
      setupPageSequence(doc, localPageSequenceProperties, sectPr);
    }

    cursor.pop();

  }

  /**
   * Handle the contents of a section
   *
   * @param doc
   * @param object
   * @return The last paragraph in the section
   * @throws DocxGenerationException
   */
  private XWPFParagraph handleSectionContent(
      XWPFDocument doc,
      XmlObject object) throws DocxGenerationException {
    XWPFParagraph lastPara = handleBody(doc, object);

    // For sections, the section properties go on the last paragraph, so if the last thing
    // in the section isn't already a paragraph, create one.
    if (lastPara == null) {
      lastPara = doc.createParagraph();
      lastPara.setSpacingBefore(0);
      lastPara.setSpacingAfter(0);
    }

    return lastPara;
  }

  /**
   * Set up page sequence properties for the entire document, including page geometry, numbering, and headers and footers.
   * @param doc Document to be constructed
   * @param xml page-sequence-properties element
   * @throws DocxGenerationException
   */
  private void setupPageSequence(XWPFDocument doc, XmlObject xml) throws DocxGenerationException {

    CTDocument1 document = doc.getDocument();
    CTBody body = (document.isSetBody() ? document.getBody() : document.addNewBody());
    CTSectPr sectPr = (body.isSetSectPr() ? body.getSectPr() : body.addNewSectPr());

    setupPageSequence(doc, xml, sectPr);
  }

  /**
   * Set up a page sequence for a section, as opposed to for the document
   * as a whole.
   * @param doc Document
   * @param object The page-sequence-properties element
   * @param sectPr The sectPr object to set the page sequence properties on.
   * @throws DocxGenerationException
   */
  private void setupPageSequence(XWPFDocument doc, XmlObject xml, CTSectPr sectPr) throws DocxGenerationException {
    XmlCursor cursor = xml.newCursor();

    setPageNumberProperties(cursor, sectPr);


    // Issue 46: Use page margins
    cursor.push();
    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-margins"))) {
      setPageMargins(cursor, sectPr);
    }
    cursor.pop();

    cursor.push();
    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "headers-and-footers"))) {
      constructHeadersAndFooters(doc, cursor.getObject(), sectPr);
    }
    cursor.pop();

    cursor.push();
    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-size"))) {
      setPageSize(cursor, sectPr);
    }
    cursor.pop();

  }

  /**
   * Set the page margins for a page sequence
   * @param cursor Cursor pointing to page-margins element
   * @param sectPr Section properties to put margins on
   */
  private void setPageMargins(XmlCursor cursor, CTSectPr sectPr) {
    CTPageMar pageMar = (sectPr.isSetPgMar() ? sectPr.getPgMar() : sectPr.addNewPgMar());
    String left = cursor.getAttributeText(DocxConstants.QNAME_LEFT_ATT);
    if (left != null) {
      try {
        long length = Measurement.toTwips(left, getDotsPerInch());
        pageMar.setLeft(BigInteger.valueOf(length));
      } catch (Exception e) {
        log.warn("setPageMargins(): Value \"" + left + " for attribute \"left\" is not a decimal number");
      }
    }
    String right = cursor.getAttributeText(DocxConstants.QNAME_RIGHT_ATT);
    if (right != null) {
      try {
        long length = Measurement.toTwips(right, getDotsPerInch());
        pageMar.setRight(BigInteger.valueOf(length));
      } catch (Exception e) {
        log.warn("setPageMargins(): Value \"" + right + " for attribute \"right\" is not a decimal number");
      }
    }
    String top = cursor.getAttributeText(DocxConstants.QNAME_TOP_ATT);
    if (top != null) {
      try {
        long length = Measurement.toTwips(top, getDotsPerInch());
        pageMar.setTop(BigInteger.valueOf(length));
      } catch (Exception e) {
        log.warn("setPageMargins(): Value \"" + top + " for attribute \"top\" is not a decimal number");
      }
    }
    String bottom = cursor.getAttributeText(DocxConstants.QNAME_BOTTOM_ATT);
    if (bottom != null) {
      try {
        long length = Measurement.toTwips(bottom, getDotsPerInch());
        pageMar.setBottom(BigInteger.valueOf(length));
      } catch (Exception e) {
        log.warn("setPageMargins(): Value \"" + bottom + " for attribute \"bottom\" is not a decimal number");
      }
    }
    String footer = cursor.getAttributeText(DocxConstants.QNAME_FOOTER_ATT);
    if (footer != null) {
      try {
        long length = Measurement.toTwips(footer, getDotsPerInch());
        pageMar.setFooter(BigInteger.valueOf(length));
      } catch (Exception e) {
        log.warn("setPageMargins(): Value \"" + footer + " for attribute \"footer\" is not a decimal number");
      }
    }
    String header = cursor.getAttributeText(DocxConstants.QNAME_HEADER_ATT);
    if (header != null) {
      try {
        long length = Measurement.toTwips(header, getDotsPerInch());
        pageMar.setHeader(BigInteger.valueOf(length));
      } catch (Exception e) {
        log.warn("setPageMargins(): Value \"" + header + " for attribute \"header\" is not a decimal number");
      }
    }
    String gutter = cursor.getAttributeText(DocxConstants.QNAME_GUTTER_ATT);
    if (gutter != null) {
      try {
        long length = Measurement.toTwips(gutter, getDotsPerInch());
        pageMar.setGutter(BigInteger.valueOf(length));
      } catch (Exception e) {
        log.warn("setPageMargins(): Value \"" + gutter + " for attribute \"gutter\" is not a decimal number");
      }
    }

  }

  private void setPageSize(XmlCursor cursor, CTSectPr sectPr) {

    CTPageSz pageSize = (sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz());
    String codeValue = cursor.getAttributeText(DocxConstants.QNAME_CODE_ATT);
    if (codeValue != null) {
      try {
        long code = Long.parseLong(codeValue);
        pageSize.setCode(BigInteger.valueOf(code));
      } catch (Exception e) {
        log.warn("setPageSize(): Value \"" + codeValue + " for attribute \"code\" is not a decimal number");
      }
    }
    String orientValue = cursor.getAttributeText(DocxConstants.QNAME_ORIENT_ATT);
    if (orientValue != null) {
      pageSize.setOrient(STPageOrientation.Enum.forString(orientValue));
    }
    String widthVal = cursor.getAttributeText(DocxConstants.QNAME_WIDTH_ATT);
    if (null != widthVal && !"".equals(widthVal.trim())) {
      try {
        long width = Measurement.toTwips(widthVal, getDotsPerInch());
        pageSize.setW(BigInteger.valueOf(width));
      } catch (MeasurementException e) {
        log.warn("setPageSize(): Value \"" + widthVal + " for attribute \"width\" cannot be converted to a twips value");
      }
    }

    String heightVal = cursor.getAttributeText(DocxConstants.QNAME_HEIGHT_ATT);
    if (null != heightVal && !"".equals(heightVal.trim())) {
      try {
        long height = Measurement.toTwips(heightVal, getDotsPerInch());
        pageSize.setH(BigInteger.valueOf(height));
      } catch (MeasurementException e) {
        log.warn("setPageSize(): Value \"" + heightVal + " for attribute \"height\" cannot be converted to a twips value");
      }
    }
  }

  private void setPageNumberProperties(XmlCursor cursor, CTSectPr sectPr) {
    cursor.push();
    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-number-properties"))) {
      String start = cursor.getAttributeText(DocxConstants.QNAME_START_ATT);
      String format = cursor.getAttributeText(DocxConstants.QNAME_FORMAT_ATT);
      String chapterSep = cursor.getAttributeText(DocxConstants.QNAME_CHAPTER_SEPARATOR_ATT);
      String chapterStyle = cursor.getAttributeText(DocxConstants.QNAME_CHAPTER_STYLE_ATT);
      if (null != format || null != chapterSep || null != chapterStyle || null != start) {
        CTPageNumber pageNumber = (sectPr.isSetPgNumType() ? sectPr.getPgNumType() : sectPr.addNewPgNumType());
        if (null != format) {
          if ("custom".equals(format)) {
            // FIXME: Implement translation from XSLT number format values to the equivalent Word
            // number formatting values.
            log.warn("Page number format \"" + format + "\" not supported. Use Word-specific values. Using \"decimal\"");
            format = "decimal";
          }
          STNumberFormat.Enum fmt = STNumberFormat.Enum.forString(format);
          if (fmt != null) {
            pageNumber.setFmt(fmt);
          }
        }
        if (chapterSep != null) {
          STChapterSep.Enum sep = STChapterSep.Enum.forString(chapterSep);
          if (sep != null) {
            pageNumber.setChapSep(sep);
          }
        }
        if (chapterStyle != null) {
          try {
            long val = Long.valueOf(chapterStyle);
            pageNumber.setChapStyle(BigInteger.valueOf(val));
          } catch (NumberFormatException e) {
            log.warn("Value \"" + chapterStyle + "\" of @chapter-style attribute is not an integer.");
          }
        }
        if (start != null) {
          try {
            long val = Long.valueOf(start);
            pageNumber.setStart(BigInteger.valueOf(val));
          } catch (NumberFormatException e) {
            log.warn("Value \"" + start + "\" of @start attribute is not an integer.");
          }
        }
      }
    }
    cursor.pop();
  }

  /**
   * Construct headers and footers on the document. If there are
   * no sections, this also sets the headers and footers for the
   * document (which acts as a single section), otherwise, each
   * section must also create the appropriate header references.
   * @param doc Document to add headers and footers to.
   * @param xml headers-and-footers element
   * @throws DocxGenerationException
   */
  private void constructHeadersAndFooters(XWPFDocument doc, XmlObject xml) throws DocxGenerationException {
    constructHeadersAndFooters(doc, xml, null);
  }

  /**
   * Construct headers and footers on the document. If there are
   * no sections, this also sets the headers and footers for the
   * document (which acts as a single section), otherwise, each
   * section must also create the appropriate header references.
   * @param doc Document to add headers and footers to.
   * @param xml headers-and-footers element
   * @param sectPr Section properties to add header and footer references to. May be null
   * @throws DocxGenerationException
   */
  private void constructHeadersAndFooters(XWPFDocument doc, XmlObject xml, CTSectPr sectPr) throws DocxGenerationException {
    XmlCursor cursor = xml.newCursor();

    boolean haveOddHeader = false;
    boolean haveEvenHeader = false;
    boolean haveOddFooter = false;
    boolean haveEvenFooter = false;

    boolean isDocument = sectPr == null;

    if (cursor.toFirstChild()) {
      XWPFHeaderFooterPolicy sectionHfPolicy = null;
      if (!isDocument) {
        sectionHfPolicy = new XWPFHeaderFooterPolicy(doc, sectPr);
      }
      do {
        String tagName = cursor.getName().getLocalPart();
        String namespace = cursor.getName().getNamespaceURI();
        List<CTHdrFtrRef> refs = null;

        if ("header".equals(tagName)) {
          HeaderFooterType type = getHeaderFooterType(cursor);
          if (type == HeaderFooterType.FIRST) {
            CTSectPr localSectPr = sectPr;
            if (localSectPr == null) {
              // FIXME: Can body be null at this time?
              localSectPr = doc.getDocument().getBody().getSectPr();
            }
            CTOnOff titlePg = (localSectPr.isSetTitlePg() ? localSectPr.getTitlePg() : localSectPr.addNewTitlePg());
            titlePg.setVal(STOnOff1.ON);
          }
          if (type == HeaderFooterType.DEFAULT) {
            haveOddHeader = true;
          }
          if (type == HeaderFooterType.EVEN) {
            haveEvenHeader = true;
          }
          if (isDocument) {
            // Make document-level header
            XWPFHeader header = doc.createHeader(type);
            makeHeaderFooter(header, cursor.getObject());
          } else {
            XWPFHeader header = sectionHfPolicy.createHeader(getSTHFTypeForXWPFHFType(type));
            makeHeaderFooter(header, cursor.getObject());
            refs = sectPr.getHeaderReferenceList();
            CTHdrFtrRef ref = getHeadeFooterRefForType(sectPr, refs, type);
            ref.setId(doc.getRelationId(header.getPart()));
            setHeaderFooterRefType(type, ref);
          }
        } else if ("footer".equals(tagName)) {
          HeaderFooterType type = getHeaderFooterType(cursor);
          if (type == HeaderFooterType.DEFAULT) {
            haveOddFooter = true;
          }
          if (type == HeaderFooterType.EVEN) {
            haveEvenFooter = true;
          }
          if (type == HeaderFooterType.FIRST) {
            CTSectPr localSectPr = sectPr;
            if (localSectPr == null) {
              // FIXME: Can body be null at this time?
              localSectPr = doc.getDocument().getBody().getSectPr();
            }
            CTOnOff titlePg = (localSectPr.isSetTitlePg() ? localSectPr.getTitlePg() : localSectPr.addNewTitlePg());
            titlePg.setVal(STOnOff1.ON);
          }
          if (isDocument) {
            // Document-level footer
            XWPFFooter footer = doc.createFooter(type);
            makeHeaderFooter(footer, cursor.getObject());
          } else {
            XWPFFooter footer = sectionHfPolicy.createFooter(getSTHFTypeForXWPFHFType(type));
            makeHeaderFooter(footer, cursor.getObject());
            refs = sectPr.getFooterReferenceList();
            CTHdrFtrRef ref = getHeadeFooterRefForType(sectPr, refs, type);
            ref.setId(doc.getRelationId(footer.getPart()));
            setHeaderFooterRefType(type, ref);
          }
        } else {
          log.warn("Unexpected element {" + namespace + "}:" + tagName + " in <headers-and-footers>. Ignored.");
        }
      } while(cursor.toNextSibling());
      if (!isDocument) {
        // setDefaultSectionHeadersAndFooters(doc, sectPr, sectionHfPolicy);
      }
      // Now set any default headers and footers from the document:
    }

    if ((haveOddHeader || haveOddFooter) &&
      (haveEvenHeader || haveEvenFooter)) {
      doc.setEvenAndOddHeadings(true);
    }

  }

  private CTHdrFtrRef getHeadeFooterRefForType(CTSectPr sectPr, List<CTHdrFtrRef> refs, HeaderFooterType type) {
    CTHdrFtrRef ref =  null;
    STHdrFtr.Enum stType = getSTHFTypeForXWPFHFType(type);
    for (CTHdrFtrRef cand : refs) {
      if (cand.getType() == stType) {
        ref = cand;
        break;
      }
    }
    if (ref == null) {
      ref = sectPr.addNewHeaderReference();
    }
    return ref;
  }

  /**
   * Sets the default headers and footers for a section, creating references to the document's
   * headers and footers, if any, for any header on the document but not already set on the section.
   * @param doc Document containing the section
   * @param sectPr Section properties for the section to set the headers on.
   * @param sectionHfPolicy The section header/footer policy that holds any headers
   * set on th esection.
   */
  public void setDefaultSectionHeadersAndFooters(
      XWPFDocument doc,
      CTSectPr sectPr,
      XWPFHeaderFooterPolicy sectionHfPolicy) {
    XWPFHeaderFooterPolicy docHfPolicy = new XWPFHeaderFooterPolicy(doc);
    if (docHfPolicy != null) {
      XWPFHeader header = null;
      XWPFFooter footer = null;
      // Default header:
      header = docHfPolicy.getDefaultHeader();
      if (sectionHfPolicy.getDefaultHeader() == null && header != null) {
        CTHdrFtrRef ref = sectPr.addNewHeaderReference();
        ref.setId(doc.getRelationId(header.getPart()));
        ref.setType(STHdrFtr.DEFAULT);
      }
      // Even header:
      header = docHfPolicy.getEvenPageHeader();
      if (sectionHfPolicy.getEvenPageHeader() == null && header != null) {
        CTHdrFtrRef ref = sectPr.addNewHeaderReference();
        ref.setId(doc.getRelationId(header.getPart()));
        ref.setType(STHdrFtr.EVEN);
      }
      // First header:
      header = docHfPolicy.getFirstPageHeader();
      if (sectionHfPolicy.getFirstPageHeader() == null && header != null) {
        CTHdrFtrRef ref = sectPr.addNewHeaderReference();
        ref.setId(doc.getRelationId(header.getPart()));
        ref.setType(STHdrFtr.FIRST);
      }
      footer = docHfPolicy.getDefaultFooter();
      if (sectionHfPolicy.getDefaultFooter() == null && footer != null) {
        CTHdrFtrRef ref = sectPr.addNewFooterReference();
        ref.setId(doc.getRelationId(footer.getPart()));
        ref.setType(STHdrFtr.DEFAULT);
      }
      // Even footer:
      footer = docHfPolicy.getEvenPageFooter();
      if (sectionHfPolicy.getEvenPageFooter() == null && footer != null) {
        CTHdrFtrRef ref = sectPr.addNewFooterReference();
        ref.setId(doc.getRelationId(footer.getPart()));
        ref.setType(STHdrFtr.EVEN);
      }
      // First footer:
      footer = docHfPolicy.getFirstPageFooter();
      if (sectionHfPolicy.getFirstPageFooter() == null && footer != null) {
        CTHdrFtrRef ref = sectPr.addNewFooterReference();
        ref.setId(doc.getRelationId(footer.getPart()));
        ref.setType(STHdrFtr.FIRST);
      }
    }
  }

  public STHdrFtr.Enum getSTHFTypeForXWPFHFType(HeaderFooterType type) {
    switch (type) {
    case EVEN:
      return STHdrFtr.EVEN;
    case FIRST:
      return STHdrFtr.FIRST;
    default:
      return STHdrFtr.DEFAULT;
    }

  }

  public void setHeaderFooterRefType(HeaderFooterType type, CTHdrFtrRef ref) {
      ref.setType(getSTHFTypeForXWPFHFType(type));
  }

  /**
   * Construct the content of a page header or footer
   * @param headerFooter {@link XPWFHeader} or {@link XWPFFooter} to add content to
   * @param xml The &lt;header&gt; or &lt;footer&gt; element to process
   * @throws DocxGenerationException
   */
  private void makeHeaderFooter(XWPFHeaderFooter headerFooter, XmlObject xml) throws DocxGenerationException {
    XmlCursor cursor = xml.newCursor();

    if (cursor.toFirstChild()) {
      do {
        String tagName = cursor.getName().getLocalPart();
        String namespace = cursor.getName().getNamespaceURI();
        if ("p".equals(tagName)) {
          XWPFParagraph p = headerFooter.createParagraph();
          makeParagraph(p, cursor);
        } else if ("table".equals(tagName)) {
          XWPFTable table = headerFooter.createTable(0, 0);
          makeTable(table, cursor.getObject());
        } else {
          // There are other body-level things that could go in a footnote but
          // we aren't worrying about them for now.
          log.warn("makeFootnote(): Unexpected element {" + namespace + "}:" + tagName + "' in <fn>. Ignored.");
        }
      } while(cursor.toNextSibling());
    }
  }

  /**
   * Get the header or footer type for the element at the cursor.
   * @param cursor
   * @return {@link HeaderFooterType}
   */
  private HeaderFooterType getHeaderFooterType(XmlCursor cursor) {
    HeaderFooterType type = HeaderFooterType.DEFAULT;
    String typeName = cursor.getAttributeText(DocxConstants.QNAME_TYPE_ATT);
    if ("even".equals(typeName)) {
      type = HeaderFooterType.EVEN;
    }
    if ("first".equals(typeName)) {
      type = HeaderFooterType.FIRST;
    }
    return type;
  }

  /**
   * Construct a Word paragraph
   * @param para The Word paragraph to construct
   * @param cursor Cursor pointing at the <p> element the paragraph will reflect.
   * @return Paragraph (should be same object as passed in).
   * @throws Exception
   */
  private XWPFParagraph makeParagraph(XWPFParagraph p, XmlCursor cursor) throws DocxGenerationException {
    return makeParagraph(p, cursor, null);
  }

  /**
   * Construct a Word paragraph
   * @param para The Word paragraph to construct
   * @param cursor Cursor pointing at the <p> element the paragraph will reflect.
   * @param additionalProperties Additional properties to add to the paragraph, i.e., from sections
   * @return Paragraph (should be same object as passed in).
   * @throws Exception
   */
  private XWPFParagraph makeParagraph(
      XWPFParagraph para,
      XmlCursor cursor,
      Map<String, String> additionalProperties)
          throws DocxGenerationException {

    cursor.push();
    String styleName = cursor.getAttributeText(DocxConstants.QNAME_STYLE_ATT);
    String styleId = cursor.getAttributeText(DocxConstants.QNAME_STYLEID_ATT);

    if (null != styleName && null == styleId) {
      // Look up the style by name:
      XWPFStyle style = para.getDocument().getStyles().getStyleWithName(styleName);
      if (null != style) {
        styleId = style.getStyleId();
      } else {
        // Try style name as styleId

        style = para.getDocument().getStyles().getStyle(styleName);
        if (null != style) {
          styleId = styleName;
        } else {

          // Issue 23: see if this is a latent style and report it
          //
          // This will require an enhancement to the POI API as there is no easy
          // way to get the list of latent styles except to parse out the XML,
          // which I'm not going to--better to fix POI.
          // Unfortunately, there does not appear to be a documented or reliable
          // way to go from Word-defined latent style names to the actual style ID
          // of the style Word *will create* by some internal magic. In addition,
          // any such mapping varies by Word version, locale, etc.
          //
          // That means that in order to use any style it must exist as a proper
          // style.
          log.warn("Paragraph style name \"" + styleName + "\" not found in the document.");
          if (this.isFirstParagraphStyleWarning) {
            // log.info("Available paragraph styles:");
            // FIXME: The POI 4.x API doesn't provide a way to get the list of styles
            //        or the list of style names that I can find short of parsing the
            //        underlying document part XML, so no way to report the
            //        style names at this time.

          }
          this.isFirstParagraphStyleWarning = false;
        }
      }
    }
    if (null != styleId) {
      para.setStyle(styleId);
    }


    if (additionalProperties != null) {
      for (String propName : additionalProperties.keySet()) {
        String value = additionalProperties.get(propName);
        if (value != null) {
          // FIXME: This is a quick hack. Need a more general
          // and elegant way to manage setting of properties.
          if (DocxConstants.PROPERTY_PAGEBREAK.equals(propName)) {
            if (DocxConstants.PROPERTY_VALUE_CONTINUOUS.equals(value)) {
              para.setPageBreak(false);
            } else {
              para.setPageBreak(true);
            }
          }
        }
      }
    }

    // Explicit page break on a paragraph should override the section-level break I would think.
    String pageBreakBefore = cursor.getAttributeText(DocxConstants.QNAME_PAGE_BREAK_BEFORE_ATT);
    if (pageBreakBefore != null) {
      boolean breakValue = Boolean.valueOf(pageBreakBefore);
      para.setPageBreak(breakValue);
    }

    // Issue 143: keepLines property.
    String keepLines = cursor.getAttributeText(DocxConstants.QNAME_KEEPLINES_ATT);
    if (keepLines != null) {
      boolean keepValue = Boolean.valueOf(keepLines);
      // If the paragraph does not already have a ctPPr then
      // we need to add one.
	  CTPPr ppr = (para.getCTP().isSetPPr() ? para.getCTP().getPPr() : para.getCTP().addNewPPr());
      CTOnOff onOff = ppr.addNewKeepLines();
      onOff.setVal(keepValue ? "on" : "off");
    }

    // Issue 143: keepNext property.
    String keepNext = cursor.getAttributeText(DocxConstants.QNAME_KEEPNEXT_ATT);
    if (keepNext != null) {
      boolean keepValue = Boolean.valueOf(keepNext);
      // If the paragraph does not already have a ctPPr then
      // we need to add one.
      @SuppressWarnings("unused")
	  CTPPr ppr = (para.getCTP().isSetPPr() ? para.getCTP().getPPr() : para.getCTP().addNewPPr());
      para.setKeepNext(keepValue);
    }

    if (cursor.toFirstChild()) {
      do {
        String tagName = cursor.getName().getLocalPart();
        String namespace = cursor.getName().getNamespaceURI();
        if ("run".equals(tagName)) {
          makeRun(para, cursor.getObject());
        } else if ("bookmarkStart".equals(tagName)) {
          makeBookmarkStart(para, cursor);
        } else if ("bookmarkEnd".equals(tagName)) {
          makeBookmarkEnd(para, cursor);
        } else if ("complexField".equals(tagName)) {
          makeComplexField(para, cursor);
        } else if ("fn".equals(tagName)) {
          makeFootnote(para, cursor.getObject());
        } else if ("hyperlink".equals(tagName)) {
          makeHyperlink(para, cursor);
        } else if ("image".equals(tagName)) {
          makeImage(para, cursor);
        } else if ("object".equals(tagName)) {
          makeObject(para, cursor);
        } else if ("page-number-ref".equals(tagName)) {
          makePageNumberRef(para, cursor);
        } else if ("math".equals(tagName) && NS_MATHML.equals(namespace)) {
          MathMLConverter.convertMath(para, cursor.getObject());
        } else {
          log.warn("Unexpected element {" + namespace + "}:" + tagName + " in <p>. Ignored.");
        }
      } while(cursor.toNextSibling());
    }
    cursor.pop();
    return para;
  }

  /**
   * Construct a page number ("PAGE") complex field.
   * @param para Paragraph to add the field to
   * @param cursor
   */
  private void makePageNumberRef(XWPFParagraph para, XmlCursor cursor) {

    String fieldData = "PAGE";
    makeSimpleField(para, fieldData);

  }

  /**
   * Makes a simple field within the specified paragraph.
   * @param para Paragraph to add the field to.
   * @param fieldData The field data, e.g. "PAGE", "DATE", etc. See 17.16 Fields and Hyperlinks.
   */
  private void makeSimpleField(XWPFParagraph para, String fieldData) {
    CTSimpleField ctField = para.getCTP().addNewFldSimple();
    ctField.setInstr(fieldData);
  }

  /**
   * Construct a run within a paragraph.
   * @param para The output paragraph to add the run to.
   * @param xml The <run> element.
   */
  private XWPFRun makeRun(XWPFParagraph para, XmlObject xml) throws DocxGenerationException {
    XmlCursor cursor = xml.newCursor();

    // if this run contains mathml then any whitespace here should *not*
    // go into the converted document
    boolean containsMathML = containsMathML(cursor);

    // String tagname = cursor.getName().getLocalPart(); // For debugging

    XWPFRun run = para.createRun();
    String styleName = cursor.getAttributeText(DocxConstants.QNAME_STYLE_ATT);
    String styleId = cursor.getAttributeText(DocxConstants.QNAME_STYLEID_ATT);

    if (null != styleName && null == styleId) {
      // Look up the style by name:
      XWPFStyle style = para.getDocument().getStyles().getStyleWithName(styleName);
      if (null != style) {
        styleId = style.getStyleId();
      } else {
        style = para.getDocument().getStyles().getStyle(styleName);
        if (null != style) {
           styleId = styleName;
        } else {
           log.warn("Character style name \"" + styleName + "\" not found in the document.");
           if (this.isFirstCharacterStyleWarning) {
             // FIXME: The POI 4.x API doesn't provide a way to get the list of styles
             //        or the list of style names that I can find short of parsing the
             //        underlying document part XML, so no way to report the
             //        style names at this time.
           }
           this.isFirstCharacterStyleWarning = false;
        }
      }
    }

    if (null != styleId) {
      run.setStyle(styleId);
    }

    handleFormattingAttributes(run, xml);

    cursor.toLastAttribute();
    cursor.toNextToken(); // Should be first text or subelement.
    // In this loop, each different token handler is responsible for positioning
    // the cursor past the thing that was handled such that the only END token
    // is the end for the run element being processed.
    while (TokenType.END != cursor.currentTokenType()) {
      // TokenType tokenType = cursor.currentTokenType(); // For debugging
      if (cursor.isText()) {
        // if this run contains mathml then any whitespace here should *not*
        // go into the converted document
        if (!containsMathML) {
          run.setText(cursor.getTextValue());
        }
        cursor.toNextToken();
      } else if (cursor.isAttr()) {
        // Ignore attributes in this context.
      } else if (cursor.isStart()) {
        // Handle element within run
        String name = cursor.getName().getLocalPart();
        String namespace = cursor.getName().getNamespaceURI();
        if ("break".equals(name)) {
          makeBreak(run, cursor);
        } else if ("symbol".equals(name)) {
          makeSymbol(run, cursor);
        } else if ("tab".equals(name)) {
          makeTab(run, cursor);
        } else if ("math".equals(name) && NS_MATHML.equals(namespace)) {
          MathMLConverter.convertMath(para, cursor.getObject());

          // converter doesn't move cursor, so we have to do it here
          cursor.toEndToken(); // now we're at </mml:math>
          // when we end, we should be on the last token before </wp:run>, but
          // there could be whitespace in between
          XmlCursor.TokenType tt = null;
          while (tt != XmlCursor.TokenType.END) {
            tt = cursor.toNextToken();
          }
          // now we are at </wp:run>, but we need to go one step back
          cursor.toPrevToken();
        } else {
          log.error("makeRun(); Unexpected element {" + namespace + "}:" + name + ". Skipping.");
          log.debug("  Element contents: " + cursor.getTextValue());
          cursor.toEndToken(); // Skip this element.
        }
        cursor.toNextToken();
      } else if (cursor.isComment() || cursor.isProcinst()) {
        // Silently ignore
        // FIXME: Not sure if we need to do more to skip a comment or processing instruction.
        cursor.toNextToken();
      } else {
        // What else could there be?
        if (cursor.getName() != null) {
          log.error("makeRun(): Unhanded XML token " + cursor.getName().getLocalPart());
        } else {
          log.error("makeRun(): Unhanded XML token " + cursor.currentTokenType());
        }
        cursor.toNextToken();
      }
    }

    cursor.pop();
    return run;
  }

  /**
   * True iff the parent element of wherever we are now contains MathML.
   */
  private boolean containsMathML(XmlCursor cursor) {
    cursor.push();

    try {
      while (cursor.currentTokenType() != TokenType.END) {
        if (cursor.currentTokenType() == TokenType.START &&
            cursor.getName().getLocalPart().equals("math") &&
            cursor.getName().getNamespaceURI().equals(NS_MATHML)) {
          return true;
        }
        cursor.toNextToken();
      }
      return false;
    } finally {
      cursor.pop();
    }
  }

  private void handleFormattingAttributes(XWPFRun run, XmlObject xml) {
    XmlCursor cursor = xml.newCursor();
    if (cursor.toFirstAttribute()) {
      do {
          String attName = cursor.getName().getLocalPart();
          String attValue = cursor.getTextValue();
          if ("bold".equals(attName)) {
              boolean value = Boolean.parseBoolean(attValue);
              run.setBold(value);
          } else if ("caps".equals(attName)) {
              boolean value = Boolean.parseBoolean(attValue);
              run.setCapitalized(value);
          } else if ("color".equals(attName)) {
            // NOTE: color must be an RGB hex number. May need to translate
            // from color names to RGB values.
              run.setColor(attValue);
          } else if ("double-strikethrough".equals(attName)) {
              boolean value = Boolean.parseBoolean(attValue);
              run.setDoubleStrikethrough(value);
          } else if ("emboss".equals(attName)) {
              boolean value = Boolean.parseBoolean(attValue);
              run.setEmbossed(value);
          } else if ("emphasis-mark".equals(attName)) {
              run.setEmphasisMark(attValue);
          } else if ("expand-collapse".equals(attName)) {
            int percentage = Integer.valueOf(attValue);
            run.setTextScale(percentage);
          } else if ("highlight".equals(attName)) {
            run.setTextHighlightColor(attValue);;
          } else if ("imprint".equals(attName)) {
              boolean value = Boolean.parseBoolean(attValue);
              run.setImprinted(value);
          } else if ("italic".equals(attName)) {
              boolean value = Boolean.parseBoolean(attValue);
              run.setItalic(value);
          } else if ("outline".equals(attName)) {
              CTOnOff onOff = CTOnOff.Factory.newInstance();
              onOff.setVal(STOnOff1.Enum.forString(attValue));
              run.getCTR().getRPr().setOutlineArray(new CTOnOff[] { onOff });
          } else if ("position".equals(attName)) {
            int val = Integer.parseInt(attValue);
              run.setTextPosition(val);
          } else if ("shadow".equals(attName)) {
              boolean value = Boolean.parseBoolean(attValue);
              run.setShadow(value);
          } else if ("small-caps".equals(attName)) {
              boolean value = Boolean.parseBoolean(attValue);
              run.setSmallCaps(value);
          } else if ("strikethrough".equals(attName)) {
              boolean value = Boolean.parseBoolean(attValue);
              run.setStrikeThrough(value);
          } else if ("underline".equals(attName)) {
              UnderlinePatterns value;
          try {
            value = UnderlinePatterns.valueOf(attValue.toUpperCase());
              run.setUnderline(value);
          } catch (Exception e) {
            log.error("- [ERROR] Unrecognized underline value \"" + attValue + "\"");
          }
          } else if ("underline-color".equals(attName)) {
            run.setUnderlineColor(attValue);
          } else if ("underline-theme-color".equals(attName)) {
            run.setUnderlineThemeColor(attValue);
          } else if ("vanish".equals(attName)) {
              boolean value = Boolean.parseBoolean(attValue);
              run.setVanish(value);
          } else if ("vertical-alignment".equals(attName)) {
              run.setVerticalAlignment(attValue);
          }

      } while(cursor.toNextAttribute());
    }

  }

  /**
   * Make a literal tabl in the run.
   * @param run
   * @param cursor
   */
  private void makeTab(XWPFRun run, XmlCursor cursor) {

    run.addTab();

  }

  /**
   * Make a symbol within a run
   * @param run
   * @param cursor
   */
  private void makeSymbol(XWPFRun run, XmlCursor cursor) {
    throw new NotImplementedException("symbol within run not implemented");

  }

  /**
   * Construct a footnote
   * @param para the paragraph containing the footnote.
   * @param cursor Pointing at the &lt;fn> element
   * @throws DocxGenerationException
   */
  private void makeFootnote(XWPFParagraph para, XmlObject xml) throws DocxGenerationException {

    XmlCursor cursor = xml.newCursor();

    String type = cursor.getAttributeText(DocxConstants.QNAME_TYPE_ATT);
    String callout = cursor.getAttributeText(DocxConstants.QNAME_CALLOUT_ATT);
    String referenceCallout = cursor.getAttributeText(DocxConstants.QNAME_REFERENCE_CALLOUT_ATT);

    XWPFAbstractFootnoteEndnote note = null;
    if ("endnote".equals(type)) {
      note = para.getDocument().createEndnote();
    } else {
      note = para.getDocument().createFootnote();
    }

    // NOTE: The footnote is not created with any initial paragraph.

    if (cursor.toFirstChild()) {
      do {
        String tagName = cursor.getName().getLocalPart();
        String namespace = cursor.getName().getNamespaceURI();
        if ("p".equals(tagName)) {
          XWPFParagraph p = note.createParagraph();
          makeParagraph(p, cursor);
        } else if ("table".equals(tagName)) {
          XWPFTable table = note.createTable();
          makeTable(table, cursor.getObject());
        } else {
          // There are other body-level things that could go in a footnote but
          // we aren't worrying about them for now.
          log.warn("makeFootnote(): Unexpected element {" + namespace + "}:" + tagName + "' in <fn>. Ignored.");
        }
      } while (cursor.toNextSibling());
    }

    para.addFootnoteReference(note);

    // Issue #29: For footnotes with explict callouts, have to replace the markup for generated
    //            refs with the literal callout from the input XML.

    if (callout != null) {
      if (referenceCallout == null) {
        referenceCallout = callout;
      }

      XmlCursor paraCursor = para.getCTP().newCursor();

      if (paraCursor.toLastChild()) {
        // Should be the run created for the footnote reference.
        if (paraCursor.toChild(DocxConstants.QNAME_FOOTNOTEREFEREMCE_ELEM)) {
          paraCursor.setAttributeText(DocxConstants.QNAME_CUSTOMMARKFOLLOWS_ATT, "on");
          paraCursor.toParent();
          paraCursor.toEndToken();
          paraCursor.insertElementWithText(DocxConstants.QNAME_T_ELEM, referenceCallout);
        }

      }

      // Set literal callout on the footnote itself:
      CTFtnEdn ctfNote = note.getCTFtnEdn();

      XmlCursor noteCursor = ctfNote.newCursor();

      // Find the first run. This should have a <w:footnoteRef/> element as it's content.
      // Remove that and replace it with a w:t with the callout.
      if (noteCursor.toChild(DocxConstants.QNAME_W_P_ELEM)) {
        if (noteCursor.toChild(DocxConstants.QNAME_R_ELEM)) {
          cursor.push();
          if (noteCursor.toChild(DocxConstants.QNAME_FOOTNOTEREF_ELEM)) {
            noteCursor.removeXml();
          }
          cursor.pop();
          // Now construct a literal footnote reference callout.
          noteCursor.insertElementWithText(DocxConstants.QNAME_T_ELEM, callout);
        }
      }

      noteCursor.close();

    }

    cursor.pop();
  }

  /**
   * Gets the current ID (i.e., the last one generated).
   * @return Current value of ID counter as a BigInteger.
   */
  @SuppressWarnings("unused")
  private BigInteger currentId() {
    return new BigInteger(Integer.toString(idCtr));
  }

  /**
   * Get the next ID for use in result objects.
   * @return Next ID value as a BitInteger
   */
  private BigInteger nextId() {
    BigInteger id = new BigInteger(Integer.toString(idCtr++));
    return id;
  }

  /**
   * Make a break within a run.
   * @param run Run to add the break to
   * @param cursor Cursor pointing to the &lt;break> element
   */
  private void makeBreak(XWPFRun run, XmlCursor cursor) throws DocxGenerationException {

    String typeValue = cursor.getAttributeText(DocxConstants.QNAME_TYPE_ATT);
    BreakType type = BreakType.TEXT_WRAPPING;
    if ("line".equals(typeValue) || "textWrapping".equals(typeValue)) {
      // Already set to this
    } else if ("page".equals(typeValue)) {
      type = BreakType.PAGE;
    } else if ("column".equals(typeValue)) {
      type = BreakType.COLUMN;
    } else {
      log.warn("makeBreak(): Unexpected @type value '" + typeValue + "'. Using 'line'.");
    }
    run.addBreak(type);
    // Now move the cursor past the end of the break element
    while (cursor.currentTokenType() != TokenType.END) {
      cursor.toNextToken();
    }
    // At this point, current token is the end of the break element
  }

  /**
   * Construct a bookmark start
   * @param para
   * @param cursor
   */
  private void makeBookmarkStart(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException
  {
    CTBookmark bookmark = para.getCTP().addNewBookmarkStart();
    String nameValue = cursor.getAttributeText(DocxConstants.QNAME_NAME_ATT);
    String idValue = cursor.getAttributeText(DocxConstants.QNAME_ID_ATT);
    if (null == nameValue || "".equals(nameValue.trim())) {
    	nameValue = idValue;
    }
    bookmark.setName(nameValue);
    BigInteger id = nextId();
    bookmark.setId(id);
    this.bookmarkIdToIdMap.put(cursor.getAttributeText(DocxConstants.QNAME_ID_ATT), id);
  }

  /**
   * Construct a bookmark end
   * @param doc
   * @param cursor
   * @throws DocxGenerationException
   */
  private void makeBookmarkEnd(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
    CTMarkupRange bookmark = para.getCTP().addNewBookmarkEnd();
    String sourceID = cursor.getAttributeText(DocxConstants.QNAME_ID_ATT);
    BigInteger id = this.bookmarkIdToIdMap.get(sourceID);
    if (id == null) {
      throw new DocxGenerationException("No bookmark start found for bookmark end with ID '" + sourceID + "'");
    } else {
      bookmark.setId(id);
    }
  }

  /**
   * Construct a complex field
   * @param para The paragraph to add the field to
   * @param cursor Points to the current XML element, which should be complexField
   */
  private void makeComplexField(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
    // Get the instruction text
    String instructionText = null;
    cursor.push();
    if (cursor.toChild(DocxConstants.QNAME_INSTRUCTIONTEXT_ELEM)) {
      instructionText = cursor.getTextValue();
    }
    cursor.pop();
    // The XWPF API has XWPFFieldRun() but no method to create one on XWPFParagraph
    // so have to construct it at the CT markup level.
    XWPFRun r = para.createRun();
    CTFldChar fldChar = r.getCTR().addNewFldChar();
    fldChar.setFldCharType(STFldCharType.BEGIN);

    // Set the instruction text
    r = para.createRun();
    CTText text = r.getCTR().addNewInstrText();
    if (instructionText == null) {
      log.warn("makeComplexField(): No instruction text for complext field: " + cursor.getTextValue());
    }
    text.setStringValue(instructionText);

    cursor.push();
    // Get the field result, if any
    if (cursor.toChild(DocxConstants.QNAME_FIELDRESULTS_ELEM)) {
      // Handle the runs
      r = para.createRun();
      fldChar = r.getCTR().addNewFldChar();
      fldChar.setFldCharType(STFldCharType.SEPARATE);
      if (cursor.toFirstChild()) {
        do {
          String tagName = cursor.getName().getLocalPart();
          String namespace = cursor.getName().getNamespaceURI();
          if ("run".equals(tagName)) {
            makeRun(para, cursor.getObject());
          } else {
            log.warn("Unexpected element {" + namespace + "}:" + tagName + " in <p>. Ignored.");
          }
        } while(cursor.toNextSibling());
      }

    }
    cursor.pop();

    r = para.createRun();
    fldChar = r.getCTR().addNewFldChar();
    fldChar.setFldCharType(STFldCharType.END);

}

  /**
   * Create runs within a hyperlink element
   * @param hyperlink The CTHyperlink to add the runs to
   * @param cursor Points to the SWPF hyperlink element
   * @param para
   * @throws Exception
   */
  private XWPFHyperlinkRun makeHyperlinkRun(
      CTHyperlink hyperlink,
      XmlCursor cursor,
      XWPFParagraph para) throws DocxGenerationException {
    // The POI 4.2 API doesn't quite match the Word structure for hyperlinks.
    // A w:hyperlink goes within a paragraph as a peer to w:run and may then contain
    // one or more w:run elements.
    //
    // The XWPF API should have XWPFHyperlinkRun implement

    // Workaround here is to add the runs to the paragraph then move
    // them to the hyperlink run.

    int runIndex =  para.getRuns().size(); // Index of first new run

    List<XWPFRun> newRuns = new ArrayList<XWPFRun>();

    // Add runs to the paragraph and capture them so we can then
    // move them to the hyperlink.
    if (cursor.toFirstChild()) {
       do {
         String tagName = cursor.getName().getLocalPart();
         String namespace = cursor.getName().getNamespaceURI();
         if ("run".equals(tagName)) {
           newRuns.add(makeRun(para, cursor.getObject()));
         } else {
           log.warn("Unexpected element {" + namespace + "}:" + tagName + " in <hyperlink>. Ignored.");
         }
      } while(cursor.toNextSibling());
    }

    XWPFHyperlinkRun hyperlinkRun = new XWPFHyperlinkRun(hyperlink, CTR.Factory.newInstance() , para);
    CTHyperlink ctHyperlink = hyperlinkRun.getCTHyperlink();

    if (newRuns.size() > 0) {
      for (XWPFRun run : newRuns) {
        CTR ctRun = run.getCTR();
        XmlCursor runCursor = ctRun.newCursor();
        if (runCursor.toFirstChild()) {
          CTR newRun = ctHyperlink.addNewR();
          // Copy the ctRun values to the new run
          do {
            String tagName = runCursor.getName().getLocalPart();
            String namespace = cursor.getName().getNamespaceURI();
            // We expect to find w:rPr and w:text
            if (tagName.equals("rPr")) {
              // Copy the rPr values
              runCursor.push();
              CTRPr newRpr = newRun.addNewRPr();
              newRpr.set(runCursor.getObject());
              runCursor.pop();
            } else if (tagName.equals("t")) {
              CTText text = newRun.addNewT();
              text.set(runCursor.getObject());
            } else {
              log.warn("Unexpected element {" + namespace + "}:" + tagName + " in run in hyperlink. Ignored.");
            }

          } while(runCursor.toNextSibling());

        }
        para.removeRun(runIndex);
      }
    } else {
      CTR run = hyperlink.addNewR();
      run.addNewT().setStringValue(cursor.getTextValue());
    }

    return hyperlinkRun;
  }

  /**
   * Construct an image reference
   * @param doc
   * @param cursor
   */
  private void makeImage(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
    cursor.push();

    // FIXME: This is all a bit scripty because of the need to get both the image data and
    //        MIME type. Not sure it's worth the effort to make it cleaner. Could define
    //        an object to hold the image data and MIME type and file name.

    String imageUrl = cursor.getAttributeText(DocxConstants.QNAME_SRC_ATT);
    if (null == imageUrl) {
      log.error("- [ERROR] No @src attribute for image.");
      return;
    }
    // Issue 72: Accept local file URLs, external URLs, and data:image/... URLs
    URI uri;
    try {
      uri = new URI(imageUrl);
    } catch (URISyntaxException e) {
      log.error("- [ERROR] " + e.getClass().getSimpleName() + " on img/@src value: " + e.getMessage());
      return;
    }
    String imageFilename = null;
    InputStream inStream;
    String mimeType = null;
    URL url;
    try {
      if (uri.isAbsolute()) {
        if ("data".equals(uri.getScheme())) {
          try {
            inStream = getStreamForDataUrl(uri);
            mimeType = getMimeTypeForDataUrl(uri);
          } catch (Exception e) {
            log.error(e.getClass().getSimpleName() + " decoding image data URL: " + e.getMessage());
            return;
          }
        } else {
          // Should be a normal URL
          url = uri.toURL();
          URLConnection conn = null;
          try {
            conn = url.openConnection();
          } catch (Exception e) {
            log.error(e.getClass().getSimpleName() + " opening image URL: " + e.getMessage());
            return;
          }
          // If we need to get the MIME type from the server, this might do it:
          // mimeType = conn.getContentEncoding();
          try {
            inStream = conn.getInputStream();
          } catch (IOException e) {
            log.error(e.getClass().getSimpleName() + " reading image URL: " + e.getMessage());
            return;
          }
          File file = new File(url.getFile());
          imageFilename = file.getName();
        }
      } else {
        // Must be relative file reference, read the file
        URL baseUrl = inFile.getParentFile().toURI().toURL();
        url = new URL(baseUrl, imageUrl);
        File file = new File(url.getFile());
        imageFilename = file.getName();
        try {
          inStream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
          log.error("Image file \"" + imageUrl + "\" not found.");
          return;
        }
      }
    } catch (MalformedURLException e) {
      log.error("- [ERROR] " + e.getClass().getSimpleName() + " on img/@src value: " + e.getMessage());
      return;
    }

    // We have to read the input stream twice, so capture the bytes
    // so we can make new streams. Can't depend on reset() on
    // the input stream (i.e., HTTP connetion stream).
    byte[] imageBytes = null;
    try {
      imageBytes = IOUtils.toByteArray(inStream);
    } catch (IOException e) {
      log.error("- [ERROR] " + e.getClass().getSimpleName() + " reading image input stream: " + e.getMessage());
      return;
    }

    // This assumes that the URL looks like a file reference.
    if (imageFilename == null || imageFilename.equals("")) {
      imageFilename = "image_" + imageCounter;
    }

    String imgExtension = FilenameUtils.getExtension(imageFilename).toLowerCase();
    int format = 0;
    if (null != imgExtension && !"".equals(imgExtension)) {
      format = getImageFormat(imgExtension);
    } else {
      format = getImageFormatForMimeType(mimeType);
    }
    int width = 200; // Default width in pixels
    int height = 200; // Default height in pixels

    if (format == 0) {
      // FIXME: Might be more appropriate to throw an exception here.
        log.error("Unsupported picture, format code \"" + format + "\": " + imageFilename +
                ". Expected emf|wmf|pict|jpeg|jpg|png|dib|gif|tiff|eps|bmp|wpg");
        cursor.pop();
        return;
    }
    BufferedImage img = null;
    int intrinsicWidth = 0;
    int intrinsicHeight = 0;
    try {
      // FIXME: Need to limit this to the formats Java2D can read.
      img = ImageIO.read(new ByteArrayInputStream(imageBytes));
      intrinsicWidth = img.getWidth();
      intrinsicHeight = img.getHeight();
    } catch (IOException e) {
      log.warn("" + e.getClass().getSimpleName() + " exception loading image file '" + imageFilename +"': " +
                     e.getMessage());
    }
    String widthVal = cursor.getAttributeText(DocxConstants.QNAME_WIDTH_ATT);
    String heightVal = cursor.getAttributeText(DocxConstants.QNAME_HEIGHT_ATT);
    boolean goodWidth = false;
    boolean goodHeight = false;

    // Issue 82: Handle empty width and height attributes (width="", height="")
    if (null != widthVal && !"".equals(widthVal.trim())) {
      try {
        width = (int) Measurement.toPixels(widthVal, getDotsPerInch());
        goodWidth = true;
      } catch (MeasurementException e) {
        log.error(e.getClass().getSimpleName() + ": " + e.getMessage());
        log.error("Using default width value " + width);
        width = intrinsicWidth > 0 ? intrinsicWidth : width;
      }
    } else {
      width = intrinsicWidth > 0 ? intrinsicWidth : width;
    }

    if (null != heightVal && !"".equals(heightVal.trim())) {
      try {
        height = (int) Measurement.toPixels(heightVal, getDotsPerInch());
        goodHeight = true;
      } catch (MeasurementException e) {
        log.error(e.getClass().getSimpleName() + ": " + e.getMessage());
        log.error("Using default height value " + height);
        height = intrinsicHeight > 0 ? intrinsicHeight : height;
      }
    } else {
      height = intrinsicHeight > 0 ? intrinsicHeight : height;
    }

    // Issue 16: If either dimension is not specified, scale the intrinsic width
    //           proportionally.
    if (widthVal == null && heightVal != null && (intrinsicWidth > 0) && goodHeight) {
      double factor = height / intrinsicHeight;
      width = (int)Math.round(intrinsicWidth * factor);
    }
    if (widthVal != null && heightVal == null && (intrinsicHeight > 0) && goodWidth) {
      double factor = (double)width / intrinsicWidth;
      height = (int)Math.round(intrinsicHeight * factor);
    }

    // At this point, the measurement is pixels. If the original specification
    // was also pixels, we need to convert to inches and then back to pixels
    // in order to apply the dots-per-inch value.

    // Word uses a DPI of 72, so if the current dotsPerInch is not 72, we need to
    // adjust the width and height by the difference.

    if (getDotsPerInch() != 72) {
      double factor = 72.0 / getDotsPerInch();
      if (widthVal != null && widthVal.matches("[0-9]+(px)?")) {
        width =  (int)Math.round(width * factor);
      }
      if (heightVal != null && heightVal.matches("[0-9]+(px)?")) {
        height = (int)Math.round(height * factor);
      }
    }


    XWPFRun run = para.createRun();

        try {
      run.addPicture(new ByteArrayInputStream(imageBytes),
                 format,
                 imageFilename,
                 Units.toEMU(width),
                 Units.toEMU(height));
    } catch (Exception e) {
      log.warn("" + e.getClass().getSimpleName() + " exception adding picture for reference '" + imageFilename +"': " +
                            e.getMessage());
    }
    imageCounter++;
    cursor.pop();
  }

  /**
   * Get the Word image format code for the specified MIME type
   * @param mimeType The MIME type to evaluate, i.e. "image/jpeg"
   * @return The format code or zero if the MIME type is not recognized.
   */
  private int getImageFormatForMimeType(String mimeType) {
    int format = 0;
    String formatString = mimeType.split("/")[1];
    format = getImageFormat(formatString);
    return format;
  }

  /**
   * Get the MIME type from a data URL.
   * @param uri The URI that is a data URL
   * @return The MIME type as a string, or null if there is no specified MIME type.
   */
  private String getMimeTypeForDataUrl(URI uri) {
    String mimeType = null;
    String url = uri.toString();
    // Data URL is data:[{mimeType}][;base64],{data}
    String[] tokens = url.substring(5).split(",");
    String props = tokens[0];
    if (props.contains(";")) {
      mimeType = props.split(";")[0];
    } else {
      if (!"".equals(props)) {
        mimeType = props;
      }
    }
    return mimeType;
  }

  /**
   * Get an input stream with the bytes from a data: URL
   * @param uri The URI that is the data: URL
   * @return Input Stream that provides access to the data bytes.
   */
  private InputStream getStreamForDataUrl(URI uri) throws Exception {
     InputStream inStream = null;
     String url = uri.toString();
     // Data URL is data:[{mimeType}][;base64],{data}
     String[] tokens = url.substring(5).split(",");
     String data = tokens[1];
     String props = tokens[0];
     if (!props.matches(".*base64")) {
       throw new Exception("data: URL does not specify \"base64\", cannot decode it. URL starts with: \"" + url.substring(0, 10));
     }
     byte[] bytes = Base64.decodeBase64(data);
     inStream = new ByteArrayInputStream(bytes);
     return inStream;
  }

  /**
   * Construct a hyperlink
   * @param doc
   * @param cursor
   * @throws Exception
   */
  private void makeHyperlink(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {

    String href = cursor.getAttributeText(DocxConstants.QNAME_HREF_ATT);

    // Hyperlink's anchor (@w:anchor) points to the name (not ID) of a bookmark.
    //
    // Alternatively, can use the @r:id attribute to point to a relationship
    // element that then points to something, normally an external resource
    // targeted by URI.

    // Convention in simple WP XML is fragment identifiers are to bookmark IDs,
    // while everything else is a URI to an external resource.

    // If the hyperlink is to a bookmark (just a fragment ID) then we we create a hyperlink with
    // an anchor, otherwise we create an external hyperlink to a URI.

    CTHyperlink hyperlink = para.getCTP().addNewHyperlink();

    // Set the appropriate target:

    if (href.startsWith("#")) {
      // Just a fragment ID, must be to a bookmark, set @anchor attribute.
      String bookmarkName = href.substring(1);
      hyperlink.setAnchor(bookmarkName);
    } else {
      // Create an external hyperlink. This creates the necessary relationship.
      // Not using the createHyperlinkRun() of paragraph because it doesn't handle the
      // runs within the <hyperlink> as we need. Set the @rId attribute.
      String rId = para.getDocument().getPackagePart().addExternalRelationship(href, XWPFRelation.HYPERLINK.getRelation()).getId();
      hyperlink.setId(rId);
    }

    cursor.push();
    XWPFHyperlinkRun hyperlinkRun = makeHyperlinkRun(hyperlink, cursor, para);
    cursor.pop();
    para.addRun(hyperlinkRun);
  }

  /**
   * Get the current dots-per-inch setting
   * @return Dots (pixels) per inch
   */
  public int getDotsPerInch() {
    return this.dotsPerInch;
  }

  /**
   * Set the dots-per-inch to use when converting from pixels to absolute measurements.
   * <p>Typical values are 72 and 96</p>
   * @param dotsPerInch The dots-per-inch value.
   */
  public void setDotsPerInch(int dotsPerInch) {
    this.dotsPerInch = dotsPerInch;
  }

  /**
   * Construct an embedded object.
   * @param para
   * @param cursor Cursor pointing to an <object> element.
   */
  private void makeObject(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
    throw new NotImplementedException("Object handling not implemented");
    //     cursor.push();
    //    cursor.pop();

  }

  /**
   * Construct an embedded object.
   * @param doc
   * @param cursor Cursor pointing to an <object> element.
   */
  private void makeObject(XWPFDocument doc, XmlCursor cursor) throws DocxGenerationException {
    throw new NotImplementedException("Object handling not implemented");
    //     cursor.push();
    //    cursor.pop();

  }

  /**
   * Construct a table.
   * @param table Table object to construct
   * @param xml The &lt;table&gt; element
   * @throws DocxGenerationException
   */
  private void makeTable(XWPFTable table, XmlObject xml) throws DocxGenerationException {

    // If the column widths are absolute measurements they can be set on the grid,
    // but if they are proportional, then they have to be set on at least the first
    // row's cells. The table grid is not required (it always reflects the calculated
    // width of the columns, possibly determined by applying percentage table and
    // column widths.
    XmlCursor cursor = xml.newCursor();

    String widthValue = cursor.getAttributeText(DocxConstants.QNAME_WIDTH_ATT);
    if (null != widthValue && !"".equals(widthValue.trim())) {
      table.setWidth(getMeasurementValue(widthValue));
    }

    setTableIndents(table, cursor);
    setTableLayout(table, cursor);



    String styleName = cursor.getAttributeText(DocxConstants.QNAME_STYLE_ATT);
    String styleId = cursor.getAttributeText(DocxConstants.QNAME_STYLEID_ATT);

    if (null != styleName && null == styleId) {
      // Look up the style by name:
      XWPFStyle style = table.getBody().getXWPFDocument().getStyles().getStyleWithName(styleName);
      if (null != style) {
        styleId = style.getStyleId();
      } else {
        log.warn("Table style name \"" + styleName + "\" not found.");
        if (this.isFirstTableStyleWarning) {
          // FIXME: The POI 4.x API doesn't provide a way to get the list of styles
          //        or the list of style names that I can find short of parsing the
          //        underlying document part XML, so no way to report the
          //        style names at this time.
        }
        this.isFirstTableStyleWarning = false;
        // Try to make a style ID out of the style name:
        styleId = styleName.replace(" ", "");
      }
    }
    if (null != styleId) {
      table.setStyleID(styleId);
    }

    TableBorderStyles borderStyles = setTableFrame(table, cursor);
    Map<QName, String> defaults = new HashMap<QName, String>();
    String rowsep = cursor.getAttributeText(DocxConstants.QNAME_ROWSEP_ATT);
    if (rowsep != null) {
      defaults.put(DocxConstants.QNAME_ROWSEP_ATT, rowsep);
    }
    String colsep = cursor.getAttributeText(DocxConstants.QNAME_COLSEP_ATT);
    if (colsep != null) {
      defaults.put(DocxConstants.QNAME_COLSEP_ATT, colsep);
    }

    int borderWidth = 8; // 8 8ths of a point, i.e. 1pt
    int borderSpace = 8; // ???
    String borderColor = "auto";

    // Rowsep is either 1 or 0
    // Default for new tables is all frames and internal borders so need
    // to explicitly set to none if rowsep or colsep is 0.
    if (rowsep != null || colsep != null) {
      if ("1".equals(rowsep)) {
        borderStyles.setRowSepBorder(borderStyles.getDefaultBorderType());
        table.setInsideHBorder(borderStyles.getRowSepBorder(), borderWidth, borderSpace, borderColor);
      } else if (rowsep != null) {
        borderStyles.setRowSepBorder(XWPFBorderType.NONE);
        table.setInsideHBorder(borderStyles.getRowSepBorder(), 0, 0, borderColor);
      }
      if ("1.".equals(colsep)) {
        borderStyles.setColSepBorder(borderStyles.getDefaultBorderType());
        table.setInsideVBorder(borderStyles.getColSepBorder(), borderWidth, borderSpace, borderColor);
      } else if (colsep !=  null) {
        borderStyles.setRowSepBorder(XWPFBorderType.NONE);
        table.setInsideVBorder(borderStyles.getRowSepBorder(), 0, 0, borderColor);
      }
    }

    // Not setting a grid on the tables because it only uses absolute
    // measurements. And there's no XWPF API for it.

    // So setting widths on columns, which allows percentages as well as
    // explicit values.
    TableColumnDefinitions colDefs = new TableColumnDefinitions();
    cursor.toChild(DocxConstants.QNAME_COLS_ELEM);
    if (cursor.toFirstChild()) {
      do {
        TableColumnDefinition colDef = colDefs.newColumnDef();

        String width = cursor.getAttributeText(DocxConstants.QNAME_COLWIDTH_ATT);
        if (null != width && !width.equals("")) {
          try {
            colDef.setWidth(width, getDotsPerInch());
          } catch (MeasurementException e) {
            log.warn("makeTable(): " + e.getClass().getSimpleName() + " - " + e.getMessage());
          }
        } else {
          colDef.setWidthAuto();
        }
      } while (cursor.toNextSibling());
    }

    // populate the rows and cells.
    cursor = xml.newCursor();

    // Header rows:
    cursor.push();
    if (cursor.toChild(DocxConstants.QNAME_THEAD_ELEM)) {
      if (cursor.toFirstChild()) {
        RowSpanManager rowSpanManager = new RowSpanManager();
        do {
          // Process the rows
          XWPFTableRow row = makeTableRow(table, cursor.getObject(), colDefs, rowSpanManager, defaults);
          row.setRepeatHeader(true);
        } while(cursor.toNextSibling());
      }
    }

    // Body rows:

    cursor = xml.newCursor();
    if (cursor.toChild(DocxConstants.QNAME_TBODY_ELEM)) {
      if (cursor.toFirstChild()) {
        RowSpanManager rowSpanManager = new RowSpanManager();
        do {
          // Process the rows
          XWPFTableRow row = makeTableRow(table, cursor.getObject(), colDefs, rowSpanManager, defaults);
          // Adjust row as needed.
          row.getCtRow(); // For setting low-level properties.
        } while(cursor.toNextSibling());
      }
    }
    table.removeRow(0); // Remove the first row that's always added automatically (FIXME: This may not be needed any more)
  }

  private void setTableIndents(XWPFTable table, XmlCursor cursor) {
    // Should only have left/right or inside/outside values, not both.

    CTTbl ctTbl = table.getCTTbl();
    CTTblPr ctTblPr = (ctTbl.getTblPr());
    if (ctTblPr == null) {
      ctTblPr = ctTbl.addNewTblPr();
    }
    String leftindentValue = cursor.getAttributeText(DocxConstants.QNAME_LEFTINDENT_ATT);
    // There only seems to be a way to set the left indent at the CT* level
//    String rightindentValue = cursor.getAttributeText(DocxConstants.QNAME_RIGHTINDENT_ATT);
//    String insideindentValue = cursor.getAttributeText(DocxConstants.QNAME_LEFTINDENT_ATT);
//    String outsideindentValue = cursor.getAttributeText(DocxConstants.QNAME_RIGHTINDENT_ATT);

    if (leftindentValue != null) {
      CTTblWidth tblWidth = CTTblWidth.Factory.newInstance();
      String value = getMeasurementValue(leftindentValue);
      try {
        tblWidth.setW(new BigInteger(value));
        ctTblPr.setTblInd(tblWidth);
      } catch (Exception e) {
        // log.debug("setTableIndents(): leftindentVale \"" + leftindentValue + "\" not an integer", e);
      }
    }

  }


  /**
   * Sets the w:tblLayout to fixed or auto
   * @param table The table object
   * @param cursor Cursor pointing at the SWPX table element.
   */
  private void setTableLayout(XWPFTable table, XmlCursor cursor) {
    // Should only have left/right or inside/outside values, not both.

    CTTbl ctTbl = table.getCTTbl();
    CTTblPr ctTblPr = (ctTbl.getTblPr());
    if (ctTblPr == null) {
      ctTblPr = ctTbl.addNewTblPr();
    }


    String layoutValue = cursor.getAttributeText(DocxConstants.QNAME_LAYOUT_ATT);

    CTTblLayoutType ctTblLayout = ctTblPr.getTblLayout();
    if (ctTblLayout == null) {
      ctTblLayout = ctTblPr.addNewTblLayout();
    }
    // #73: Make default layout be "autofit" to match behavior before
    //      implementation of #49
    ctTblLayout.setType(STTblLayoutType.AUTOFIT);
    if (layoutValue != null && !"auto".equals(layoutValue)) {
      ctTblLayout.setType(STTblLayoutType.FIXED);
    }

  }

  /**
   * Get the word measurement value as either a keyword, a percentage, or a twips integer.
   * @param measurement The measurement to convert
   * @return Twips value, percentage, or "auto"
   */
  public String getMeasurementValue(String measurement) {
    String result = "auto";
    if (measurement.endsWith("%") || measurement.equals("auto")) {
      result = measurement;
    } else {
      try {
        long twips = Measurement.toTwips(measurement, getDotsPerInch());
        result = "" + twips;
      } catch (Exception e) {
        log.warn("getMeasurementValue(): " + e.getClass().getSimpleName() + " - " + e.getMessage(), e);
      }
    }
    return result;
  }

  private TableBorderStyles setTableFrame(XWPFTable table, XmlCursor cursor) {
    int frameWidth = 8; // 1pt
    int frameSpace = 0;
    String frameColor = "auto";

    String frameValue = cursor.getAttributeText(DocxConstants.QNAME_FRAME_ATT);

    TableBorderStyles borderStyles =
        new TableBorderStyles(cursor.getObject());

    XWPFBorderType topBorder = borderStyles.getTopBorder();
    XWPFBorderType bottomBorder = borderStyles.getBottomBorder();
    XWPFBorderType leftBorder = borderStyles.getLeftBorder();
    XWPFBorderType rightBorder = borderStyles.getRightBorder();


    if (frameValue != null) {
      if ("none".equals(frameValue)) {
        topBorder = XWPFBorderType.NONE;
        bottomBorder = XWPFBorderType.NONE;
        leftBorder = XWPFBorderType.NONE;
        rightBorder = XWPFBorderType.NONE;
      } else if ("all".equals(frameValue)) {
        topBorder = getBorderStyle(topBorder, borderStyles.getDefaultBorderType());
        bottomBorder = getBorderStyle(bottomBorder, borderStyles.getDefaultBorderType());
        leftBorder = getBorderStyle(leftBorder, borderStyles.getDefaultBorderType());
        rightBorder = getBorderStyle(rightBorder, borderStyles.getDefaultBorderType());
      } else if ("topbot".equals(frameValue)) {
        topBorder = getBorderStyle(topBorder, borderStyles.getDefaultBorderType());
        bottomBorder = getBorderStyle(bottomBorder, borderStyles.getDefaultBorderType());
        leftBorder = XWPFBorderType.NONE;
        rightBorder = XWPFBorderType.NONE;
      } else if ("sides".equals(frameValue)) {
        topBorder = XWPFBorderType.NONE;
        bottomBorder = XWPFBorderType.NONE;
        leftBorder = getBorderStyle(leftBorder, borderStyles.getDefaultBorderType());
        rightBorder = getBorderStyle(rightBorder, borderStyles.getDefaultBorderType());
      } else if ("top".equals(frameValue)) {
        topBorder = getBorderStyle(topBorder, borderStyles.getDefaultBorderType());
        bottomBorder = XWPFBorderType.NONE;
        leftBorder = XWPFBorderType.NONE;
        rightBorder = XWPFBorderType.NONE;
      } else if ("bottom".equals(frameValue)) {
        topBorder = XWPFBorderType.NONE;
        bottomBorder = getBorderStyle(bottomBorder, borderStyles.getDefaultBorderType());
        leftBorder = XWPFBorderType.NONE;
        rightBorder = XWPFBorderType.NONE;
      }

    }
    if (bottomBorder != null) {
      table.setBottomBorder(bottomBorder, frameWidth, frameSpace, frameColor);
    }
    if (topBorder != null) {
      table.setTopBorder(topBorder, frameWidth, frameSpace, frameColor);
    }
    if (leftBorder != null) {
      table.setLeftBorder(leftBorder, frameWidth, frameSpace, frameColor);
    }
    if (rightBorder != null) {
      table.setRightBorder(rightBorder, frameWidth, frameSpace, frameColor);
    }
    return borderStyles;
  }

  /**
   * Get the border style, using the default if the explicit style null
   * @param explicitStyle Explicitly-specified border style. May be null
   * @param defaultType The default to use if explicit is null
   * @return The effective border style
   */
  private XWPFBorderType getBorderStyle(XWPFBorderType explictType, XWPFBorderType defaultType) {
    return (explictType == null ? defaultType : explictType);
  }

  /**
   * Get the XWPFBorderType for the specified STBorder value.
   * @param borderValue Border value (e.g., "wave").
   * @return Corresponding XWPFBorderType value or null if there is no corresponding value.
   */
  private XWPFBorderType xwpfBorderType(String borderValue) {

    STBorder.Enum borderStyle = STBorder.Enum.forString(borderValue);

    // There's not a direct correspondence between STBorder int values
    // and XWPFBorderType so just building a switch statement.
    XWPFBorderType xwpfType = null;
    switch (borderStyle.intValue()) {
    case STBorder.INT_DOT_DASH:
      xwpfType = XWPFBorderType.DOT_DASH;
      break;
    case STBorder.INT_DASH_SMALL_GAP:
      xwpfType = XWPFBorderType.DASH_SMALL_GAP;
      break;
    case STBorder.INT_DASH_DOT_STROKED:
      xwpfType = XWPFBorderType.DASH_DOT_STROKED;
      break;
    case STBorder.INT_DASHED:
      xwpfType = XWPFBorderType.DASHED;
      break;
    case STBorder.INT_DOT_DOT_DASH:
      xwpfType = XWPFBorderType.DOT_DOT_DASH;
      break;
    case STBorder.INT_DOTTED:
      xwpfType = XWPFBorderType.DOTTED;
      break;
    case STBorder.INT_DOUBLE:
      xwpfType = XWPFBorderType.DOUBLE;
      break;
    case STBorder.INT_DOUBLE_WAVE:
      xwpfType = XWPFBorderType.DOUBLE_WAVE;
      break;
    case STBorder.INT_INSET:
      xwpfType = XWPFBorderType.INSET;
      break;
    case STBorder.INT_NIL:
      xwpfType = XWPFBorderType.NIL;
      break;
    case STBorder.INT_NONE:
      xwpfType = XWPFBorderType.NONE;
      break;
    case STBorder.INT_OUTSET:
      xwpfType = XWPFBorderType.OUTSET;
      break;
    case STBorder.INT_SINGLE:
      xwpfType = XWPFBorderType.SINGLE;
      break;
    case STBorder.INT_THICK:
      xwpfType = XWPFBorderType.THICK;
      break;
    case STBorder.INT_THICK_THIN_LARGE_GAP:
      xwpfType = XWPFBorderType.THICK_THIN_LARGE_GAP;
      break;
    case STBorder.INT_THICK_THIN_MEDIUM_GAP:
      xwpfType = XWPFBorderType.THICK_THIN_MEDIUM_GAP;
      break;
    case STBorder.INT_THICK_THIN_SMALL_GAP:
      xwpfType = XWPFBorderType.THICK_THIN_SMALL_GAP;
      break;
    case STBorder.INT_THIN_THICK_LARGE_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_LARGE_GAP;
      break;
    case STBorder.INT_THIN_THICK_MEDIUM_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_MEDIUM_GAP;
      break;
    case STBorder.INT_THIN_THICK_SMALL_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_SMALL_GAP;
      break;
    case STBorder.INT_THIN_THICK_THIN_LARGE_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_THIN_LARGE_GAP;
      break;
    case STBorder.INT_THIN_THICK_THIN_MEDIUM_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_THIN_MEDIUM_GAP;
      break;
    case STBorder.INT_THIN_THICK_THIN_SMALL_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_THIN_SMALL_GAP;
      break;
    case STBorder.INT_THREE_D_EMBOSS:
      xwpfType = XWPFBorderType.THREE_D_EMBOSS;
      break;
    case STBorder.INT_THREE_D_ENGRAVE:
      xwpfType = XWPFBorderType.THREE_D_ENGRAVE;
      break;
    case STBorder.INT_TRIPLE:
      xwpfType = XWPFBorderType.TRIPLE;
      break;
    case STBorder.INT_WAVE:
      xwpfType = XWPFBorderType.WAVE;
      break;
    }
    return xwpfType;
  }

  /**
   * Get the STBorderType.Enum for the specified STBorder value.
   * @param borderValue Border value (e.g., "wave").
   * @return Corresponding XWPFBorderType value or null if there is no corresponding value.
   */
  private STBorder.Enum stBorderType(XWPFBorderType borderType) {

    // There's not a direct correspondence between STBorder int values
    // and XWPFBorderType so just building a switch statement.
    STBorder.Enum stBorder = null;
    switch (borderType) {
    case DOT_DASH:
      stBorder = STBorder.DOT_DASH;
      break;
    case DASH_SMALL_GAP:
      stBorder = STBorder.DASH_SMALL_GAP;
      break;
    case DASH_DOT_STROKED:
      stBorder = STBorder.DASH_DOT_STROKED;
      break;
    case DASHED:
      stBorder = STBorder.DASHED;
      break;
    case DOT_DOT_DASH:
      stBorder = STBorder.DOT_DOT_DASH;
      break;
    case DOTTED:
      stBorder = STBorder.DOTTED;
      break;
    case DOUBLE:
      stBorder = STBorder.DOUBLE;
      break;
    case DOUBLE_WAVE:
      stBorder = STBorder.DOUBLE_WAVE;
      break;
    case INSET:
      stBorder = STBorder.INSET;
      break;
    case NIL:
      stBorder = STBorder.NIL;
      break;
    case NONE:
      stBorder = STBorder.NONE;
      break;
    case OUTSET:
      stBorder = STBorder.OUTSET;
      break;
    case SINGLE:
      stBorder = STBorder.SINGLE;
      break;
    case THICK:
      stBorder = STBorder.THICK;
      break;
    case THICK_THIN_LARGE_GAP:
      stBorder = STBorder.THICK_THIN_LARGE_GAP;
      break;
    case THICK_THIN_MEDIUM_GAP:
      stBorder = STBorder.THICK_THIN_MEDIUM_GAP;
      break;
    case THICK_THIN_SMALL_GAP:
      stBorder = STBorder.THICK_THIN_SMALL_GAP;
      break;
    case THIN_THICK_LARGE_GAP:
      stBorder = STBorder.THIN_THICK_LARGE_GAP;
      break;
    case THIN_THICK_MEDIUM_GAP:
      stBorder = STBorder.THIN_THICK_MEDIUM_GAP;
      break;
    case THIN_THICK_SMALL_GAP:
      stBorder = STBorder.THIN_THICK_SMALL_GAP;
      break;
    case THIN_THICK_THIN_LARGE_GAP:
      stBorder = STBorder.THIN_THICK_THIN_LARGE_GAP;
      break;
    case THIN_THICK_THIN_MEDIUM_GAP:
      stBorder = STBorder.THIN_THICK_THIN_MEDIUM_GAP;
      break;
    case THIN_THICK_THIN_SMALL_GAP:
      stBorder = STBorder.THIN_THICK_THIN_SMALL_GAP;
      break;
    case THREE_D_EMBOSS:
      stBorder = STBorder.THREE_D_EMBOSS;
      break;
    case THREE_D_ENGRAVE:
      stBorder = STBorder.THREE_D_ENGRAVE;
      break;
    case TRIPLE:
      stBorder = STBorder.TRIPLE;
      break;
    case WAVE:
      stBorder = STBorder.WAVE;
      break;
    }
    return stBorder;
  }

  /**
   * Construct a table row
   * @param table The table to add the row to
   * @param xml The <row> element to add to the table
   * @param colDefs Column definitions
   * @param rowSpanManager Manages setting vertical spanning across multiple rows.
   * @param defaults Defaults inherited from the table (or elsewhere)
   * @return Constructed row object
   * @throws DocxGenerationException
   */
  private XWPFTableRow makeTableRow(
      XWPFTable table,
      XmlObject xml,
      TableColumnDefinitions colDefs,
      RowSpanManager rowSpanManager,
      Map<QName, String> defaults)
          throws DocxGenerationException {
    XmlCursor cursor = xml.newCursor();
    XWPFTableRow row = table.createRow();

    cursor.push();
    cursor.toChild(DocxConstants.QNAME_TD_ELEM);
    int cellCtr = 0;

    do {
      // log.debug("makeTableRow(): Cell " + cellCtr);
      TableColumnDefinition colDef = colDefs.get(cellCtr);
      // Rows always have at least one cell
      // FIXME: At some point the POI API will remove the automatic creation
      // of the first cell in a row.
      XWPFTableCell cell = cellCtr == 0 ? row.getCell(0) : row.addNewTableCell();

      CTTcPr ctTcPr = cell.getCTTc().addNewTcPr();
      String align = cursor.getAttributeText(DocxConstants.QNAME_ALIGN_ATT);
      String valign = cursor.getAttributeText(DocxConstants.QNAME_VALIGN_ATT);
      String colspan = cursor.getAttributeText(DocxConstants.QNAME_COLSPAN_ATT);
      String rowspan = cursor.getAttributeText(DocxConstants.QNAME_ROWSPAN_ATT);
      String shade = cursor.getAttributeText(DocxConstants.QNAME_SHADE_ATT);

      setCellBorders(cursor, ctTcPr);
      long spanCount = 1; // Default value;

      try {
        String widthValue = cursor.getAttributeText(DocxConstants.QNAME_WIDTH_ATT);
        if (null != widthValue && !"".equals(widthValue.trim())) {
          cell.setWidth(TableColumnDefinition.interpretWidthSpecification(widthValue, getDotsPerInch()));
        } else {
          String width = null;
          width = colDef.getWidth();
          if (colspan != null) {
            try {
              spanCount = Integer.parseInt(colspan);
              // Try to add up the widths of the spanned columns.
              // This is only possible if the values are all percents
              // or are all measurements. Since we don't the actual
              // width of the table itself necessarily, there's no way
              // to reliably convert percentages to explicit widths.
              List<String> spanWidths = new ArrayList<String>();
              boolean allPercents = true;
              boolean allNumbers = true;
              boolean allAuto = true;
              for (int i = cellCtr; i < cellCtr + spanCount; i++) {
                String widthVal = colDefs.get(i).getSpecifiedWidth();
                spanWidths.add(widthVal);
                allPercents = allPercents && widthVal.endsWith("%");
                allNumbers = allNumbers && !widthVal.endsWith("%") && !widthVal.equals("auto");
                allAuto = allAuto && widthVal.equals("auto");
              }
              if (allPercents) {
                double spanPercent = 0;
                for (String cand : spanWidths) {
                  String number = cand.substring(0, cand.lastIndexOf("%"));
                  try {
                    spanPercent += Double.parseDouble(number);
                  } catch (NumberFormatException e) {
                    log.warn("Calculating width of column-spanning cell: Expected percent value \"" + cand + "\" is not numeric.");
                  }
                }
                width = "" + spanPercent + "%";
              } else if (allAuto) {
                // Set widths to equal percents so we can calculate span widths.
                int colCount = colDefs.getColumnDefinitions().size();
                double spanPercent = 100.0 / colCount;
                width = "" + spanPercent + "%";
              } else if (allNumbers) {
                int spanMeasurement = 0;
                for (String cand : spanWidths) {
                  String number = TableColumnDefinition.interpretWidthSpecification(cand, getDotsPerInch());
                  try {
                    spanMeasurement += Integer.parseInt(number);
                  } catch (NumberFormatException e) {
                    log.warn("Expected percent value \"" + cand + "\" is not numeric.");
                  }
                }
                width = "" + spanMeasurement;
              } else {
                log.warn("Widths of spanned columns are neither all percents, all auto, or all measurements, cannot calculate exact spanned width");
                log.warn("Widths are \"" + String.join("\", \"", spanWidths) + "\"");
              }
              cell.setWidth(width);
            } catch (Exception e) {
              log.error("makeTableRow(): @colspan value \"" + colspan + "\" is not an integer. Using first column's width.");
            }

          }
          cell.setWidth(width);
          //log.debug("makeTableRow():   Setting width from column definition: " + colDef.getWidth() + " (" + colDef.getSpecifiedWidth() + ")");
        }
      } catch (Exception e) {
        log.error(e.getClass().getSimpleName() + " setting width for column " + (cellCtr + 1) + ": " + e.getMessage(), e);
      }
      if (null != valign) {
        XWPFVertAlign vertAlign = XWPFVertAlign.valueOf(valign.toUpperCase());
        cell.setVerticalAlignment(vertAlign);
      }
      if (null != colspan) {
        try {
          int spanval = Integer.parseInt(colspan);
          CTDecimalNumber spanNumber = CTDecimalNumber.Factory.newInstance();
          spanNumber.setVal(BigInteger.valueOf(spanval));
          // Set the gridspan on the cell to the span count. This will usually
          // set up the width correctly when Word lays out the table
          // regardless of what the nominal column width is. This is because
          // Word infers the table grid from the columns and cells automatically.
          // However, it appears this doesn't always work as expected.
          ctTcPr.setGridSpan(spanNumber);
        } catch (NumberFormatException e) {
          log.warn("Non-numeric value for @colspan: \"" + colspan + "\". Ignored.");
        }
      }
      if (null != rowspan) {
        try {
          int spanval = Integer.parseInt(rowspan);
          CTDecimalNumber spanNumber = CTDecimalNumber.Factory.newInstance();
          spanNumber.setVal(BigInteger.valueOf(spanval));
          rowSpanManager.addColumn(cellCtr, spanval);
          CTVMerge vMerge = CTVMerge.Factory.newInstance();
          vMerge.setVal(STMerge.RESTART);
          ctTcPr.setVMerge(vMerge);
        } catch (NumberFormatException e) {
          log.warn("Non-numeric value for @rowspan: \"" + rowspan + "\". Ignored.");
        }
      }

      if (null != shade) {
        try {
          CTShd ctShd = CTShd.Factory.newInstance();
          // <w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>
          ctShd.setFill(shade);
          ctShd.setColor("auto");
          ctShd.setVal(STShd.CLEAR);
          ctTcPr.setShd(ctShd);
        } catch (Exception e) {
          log.warn("Shade value must be a 6-digit hex string, got \"" + shade + "\"");
        }
      }

      cursor.push();
      // The first cell of a span will already have a vertical span set for it.
      if (rowspan == null && cursor.toChild(DocxConstants.QNAME_VSPAN_ELEM)) {
        int spansRemaining = rowSpanManager.includeCell(cellCtr);
        if (spansRemaining < 0) {
          log.warn("Found <vspan> when there should not have been one. Ignored.");
        } else {
          ctTcPr.setVMerge(CTVMerge.Factory.newInstance());
        }
      } else {
        // convert the contents of the cell
        boolean hasMore = cursor.toFirstChild();
        // Issue 134: If <td> is empty, hasMore will be false.
        if (!hasMore) {
          // Leave the empty paragraph, which is required by Word.
        } else {
          // Cells always have at least one paragraph.
          cell.removeParagraph(0);

          while (hasMore) {
            if (cursor.getName().equals(DocxConstants.QNAME_P_ELEM)) {
              XWPFParagraph p = cell.addParagraph();
              makeParagraph(p, cursor);
              if (null != align) {
                if ("JUSTIFY".equalsIgnoreCase(align)) {
                  // Issue 18: "BOTH" is the better match to "JUSTIFY"
                  align = "BOTH"; // Slight mistmatch between markup and model
                }
                if ("CHAR".equalsIgnoreCase(align)) {
                  // I'm not sure this is the best mapping but it seemed close enough
                  align = "NUM_TAB"; // Slight mistmatch between markup and model
                }
                ParagraphAlignment alignment = ParagraphAlignment.valueOf(align.toUpperCase());
                p.setAlignment(alignment);
              }
            } else if (cursor.getName().equals(DocxConstants.QNAME_TABLE_ELEM)) {
              // record how many tables were in the cell previously
              int preTables = cell.getCTTc().getTblList().size();
  
              CTTbl ctTbl = cell.getCTTc().addNewTbl();
              ctTbl = cell.getCTTc().addNewTbl();
              CTTblPr tblPr = ctTbl.addNewTblPr();
              tblPr.addNewTblW();
  
              XWPFTable nestedTable = new XWPFTable(ctTbl, cell);
              makeTable(nestedTable, cursor.getObject());
  
              // for some reason this inserts two tables, where the
              // first one is empty. we need to remove that one.
              // luckily, the number of tables we used to have equals
              // the index of the first new table
              cell.getCTTc().removeTbl(preTables);
            } else {
              log.warn("Table cell contains unknown element {} -- skipping", cursor.getName());
            }
  
            hasMore = cursor.toNextSibling();
          }
        }
      }
      cursor.pop();
      cellCtr += spanCount;
    } while(cursor.toNextSibling());
    return row;
  }


  /**
   * Set the borders on the cells.
   * @param cursor cursor for the table cell markup
   * @param ctTcPr Table cell style properties
   * @return
   */
  private void setCellBorders(XmlCursor cursor, CTTcPr ctTcPr) {

    // log.debug("setCellBorders(): tag is \"" + cursor.getName().getLocalPart() + "\"");
    TableBorderStyles borderStyles = new TableBorderStyles(cursor.getObject());

    if (borderStyles.hasBorders()) {
      CTTcBorders borders = ctTcPr.addNewTcBorders();

      // Borders can be set per edge:

      if (borderStyles.getBottomBorder() != null) {
        CTBorder bottom = borders.addNewBottom();
        STBorder.Enum val = borderStyles.getBottomBorderEnum();
        if (val != null) {
          bottom.setVal(val);
        } else {
          log.warn("setCellBorders(): Failed to get STBorder.Enum value for XWPFBorderStyle \"" + borderStyles.getBottomBorder().name() + "\"");
        }
        if (borderStyles.getBottomColor() != null) {
          bottom.setColor(borderStyles.getBottomColor());
        }
      }
      if (borderStyles.getTopBorder() != null) {
        CTBorder top = borders.addNewTop();
        top.setVal(borderStyles.getTopBorderEnum());
        if (borderStyles.getTopColor() != null) {
          top.setColor(borderStyles.getTopColor());
        }
      }
      if (borderStyles.getLeftBorder() != null) {
        CTBorder left = borders.addNewLeft();
        left.setVal(borderStyles.getLeftBorderEnum());
        if (borderStyles.getLeftColor() != null) {
          left.setColor(borderStyles.getLeftColor());
        }
      }
      if (borderStyles.getRightBorder() != null) {
        CTBorder right = borders.addNewRight();
        right.setVal(borderStyles.getRightBorderEnum());
        if (borderStyles.getRightColor() != null) {
          right.setColor(borderStyles.getRightColor());
        }
      }
    }
  }

  private void setupStyles(XWPFDocument doc, XWPFDocument templateDoc) throws DocxGenerationException {
    // Load template. For now this is hard coded but will need to be
    // parameterized

    // Copy the template's styles to result document:

    try {
      XWPFStyles newStyles = doc.createStyles();
      newStyles.setStyles(templateDoc.getStyle());
    } catch (IOException e) {
      new DocxGenerationException(e.getClass().getSimpleName() + " reading template DOCX file: " + e.getMessage(), e);
    } catch (XmlException e) {
      new DocxGenerationException(e.getClass().getSimpleName() + " Copying styles from template doc: " + e.getMessage(), e);
    }


  }

  private void setupNumbering(XWPFDocument doc, XWPFDocument templateDoc) throws DocxGenerationException {
    // Load the template's numbering definitions to the new document

    try {
      XWPFNumbering templateNumbering = templateDoc.getNumbering();
      XWPFNumbering numbering = doc.createNumbering();
      // In 4.1.2 There is no method to just get all the abstract and concrete
      // numbers or their IDs so we just iterate until we don't get any more
      // Trunk has new methods for this as of 4/26/2020

      // Abstract numbers:
      int i = 1;

      XWPFAbstractNum abstractNum = null;
      // Number IDs appear to always be integers starting at 1
      // so we're really just guessing.
      do {
        abstractNum = templateNumbering.getAbstractNum(BigInteger.valueOf(i));
        i++;
        if (abstractNum != null) {
          numbering.addAbstractNum(abstractNum);
        }
      } while (abstractNum != null);

      // Concrete numbers:
      XWPFNum num = null;
      i = 1;
      do {
        num = templateNumbering.getNum(BigInteger.valueOf(i));
        i++;
        if (num != null) {
          numbering.addNum(num);
        }
      } while (num != null);


    } catch (Exception e) {
      new DocxGenerationException(e.getClass().getSimpleName() + " Copying numbering definitions from template doc: " + e.getMessage(), e);
    }


  }

  /**
   * Set up any custom styles.
   * @param doc Word doc to set up styles for
   */
  @SuppressWarnings("unused")
  private void setupFootnoteStyles(XWPFDocument doc) throws DocxGenerationException {

    // Styles for footnotes:

    doc.createStyles(); // Make sure we have styles

    CTStyle style = CTStyle.Factory.newInstance();
    style.setStyleId("FootnoteReference");
    style.setType(STStyleType.CHARACTER);
    style.addNewName().setVal("footnote reference");
    style.addNewBasedOn().setVal("DefaultParagraphFont");
    style.addNewUiPriority().setVal(new BigInteger("99"));
    style.addNewSemiHidden();
    style.addNewUnhideWhenUsed();
    style.addNewRPr().addNewVertAlign().setVal(STVerticalAlignRun.SUPERSCRIPT);

    doc.getStyles().addStyle(new XWPFStyle(style));

    style = CTStyle.Factory.newInstance();
    style.setType(STStyleType.PARAGRAPH);
    style.setStyleId("FootnoteText");
    style.addNewName().setVal("footnote text");
    style.addNewBasedOn().setVal("Normal");
    style.addNewLink().setVal("FootnoteTextChar");
    style.addNewUiPriority().setVal(new BigInteger("99"));
    style.addNewSemiHidden();
    style.addNewUnhideWhenUsed();
    CTRPr rpr = style.addNewRPr();
    rpr.addNewSz().setVal(new BigInteger("20"));
    rpr.addNewSzCs().setVal(new BigInteger("20"));

    doc.getStyles().addStyle(new XWPFStyle(style));

    style  = CTStyle.Factory.newInstance();
    style.setCustomStyle(STOnOff1.OFF);
    style.setStyleId("FootnoteTextChar");
    style.setType(STStyleType.CHARACTER);
    style.addNewName().setVal("Footnote Text Char");
    style.addNewBasedOn().setVal("DefaultParagraphFont");
    style.addNewLink().setVal("FootnoteText");
    style.addNewUiPriority().setVal(new BigInteger("99"));
    style.addNewSemiHidden();
    rpr = style.addNewRPr();
    rpr.addNewSz().setVal(new BigInteger("20"));
    rpr.addNewSzCs().setVal(new BigInteger("20"));

    doc.getStyles().addStyle(new XWPFStyle(style));

  }

  /**
   * Get the Word-specific format value.
   * @param imgExtension
   * @return The format or 0 (zero) if the format is not recognized.
   */
  private int getImageFormat(String imgExtension) {
    int format = 0;

    if ("emf".equals(imgExtension)) format = XWPFDocument.PICTURE_TYPE_EMF;
        else if ("wmf".equals(imgExtension)) format = XWPFDocument.PICTURE_TYPE_WMF;
        else if ("pict".equals(imgExtension)) format = XWPFDocument.PICTURE_TYPE_PICT;
        else if ("jpeg".equals(imgExtension) ||
             "jpg".equals(imgExtension)) format = XWPFDocument.PICTURE_TYPE_JPEG;
        else if ("png".equals(imgExtension)) format = XWPFDocument.PICTURE_TYPE_PNG;
        else if ("dib".equals(imgExtension)) format = XWPFDocument.PICTURE_TYPE_DIB;
        else if ("gif".equals(imgExtension)) format = XWPFDocument.PICTURE_TYPE_GIF;
        else if ("tiff".equals(imgExtension)) format = XWPFDocument.PICTURE_TYPE_TIFF;
        else if ("eps".equals(imgExtension)) format = XWPFDocument.PICTURE_TYPE_EPS;
        else if ("bmp".equals(imgExtension)) format = XWPFDocument.PICTURE_TYPE_BMP;
        else if ("wpg".equals(imgExtension)) format = XWPFDocument.PICTURE_TYPE_WPG;

    return format;
  }


}
