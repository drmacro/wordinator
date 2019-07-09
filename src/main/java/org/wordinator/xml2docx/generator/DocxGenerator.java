/**
 * 
 */
package org.wordinator.xml2docx.generator;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;
import javax.xml.namespace.QName;

import org.apache.commons.io.FilenameUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFAbstractFootnoteEndnote;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFHeaderFooter;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtrRef;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STChapterSep;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STSectionMark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalAlignRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.STOnOffImpl;
import org.wordinator.xml2docx.xwpf.model.XWPFHeaderFooterPolicy;

/**
 * Generates DOCX files from Simple Word Processing Markup Language XML.
 */
public class DocxGenerator {
	
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

  public static final Logger log = LogManager.getLogger();

	private File outFile;
	private int dotsPerInch = 72; /* DPI */
	// Map of source IDs to internal object IDs.
	private Map<String, BigInteger> bookmarkIdToIdMap = new HashMap<String, BigInteger>();
	private int idCtr = 0;
	private File inFile;
	private XWPFDocument templateDoc;

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
		XmlObject pageSequenceProperties = null;
		if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-sequence-properties"))) {
		  // Set up document-level headers. These will apply to the whole
		  // document if there are no sections, or to the last section if
		  // there are sections. Results in a w:sectPr as  the last child 
		  // of w:body.
      setupPageSequence(doc, cursor.getObject());
			pageSequenceProperties = cursor.getObject();
		}
		cursor.pop();
		cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "body"));
		handleBody(doc, cursor.getObject(), pageSequenceProperties);
		
		
	}

  /**
	 * Process the elements in &lt;body&gt;
	 * @param doc Document to add paragraphs to.
	 * @param xml Body element
   * @param pageSequenceProperties Document-level page sequence properties. Used
   * if there are no section-level page sequence properties.
   * @return Last paragraph of the body (if any)
	 * @throws DocxGenerationException
	 */
	private XWPFParagraph handleBody(
	    XWPFDocument doc, 
	    XmlObject xml, XmlObject pageSequenceProperties) 
	        throws DocxGenerationException {
	  if (log.isDebugEnabled()) {
	    // log.debug("handleBody(): starting...");
	  }
		XmlCursor cursor = xml.newCursor();
		if (cursor.toFirstChild()) {
			do {
				String tagName = cursor.getName().getLocalPart();
				String namespace = cursor.getName().getNamespaceURI();
				if ("p".equals(tagName)) {
					XWPFParagraph p = doc.createParagraph();
					makeParagraph(p, cursor);
				} else if ("section".equals(tagName)) {
					handleSection(doc, cursor.getObject(), pageSequenceProperties);
				} else if ("table".equals(tagName)) {
					XWPFTable table = doc.createTable();
					makeTable(table, cursor.getObject());
				} else if ("object".equals(tagName)) {
					// FIXME: This is currently unimplemented.
					makeObject(doc, cursor);
				} else {
					log.warn("handleBody(): Unexpected element {" + namespace + "}:'" + tagName + "' in <body>. Ignored.");
				}
			} while (cursor.toNextSibling());
			
		}
    // The section properties always go on an empty paragraph.
		XWPFParagraph lastPara = doc.createParagraph();
    lastPara.setSpacingBefore(0);
    lastPara.setSpacingAfter(0);
		return lastPara;
	}

	/**
	 * Handle a &lt;section&gt; element
	 * @param doc Document we're adding to
	 * @param xml &lt;section&gt; element
	 * @param docPageSequenceProperties Document-level page sequence properties
	 */
	private void handleSection(
	    XWPFDocument doc, 
	    XmlObject xml, 
	    XmlObject docPageSequenceProperties) 
	        throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();
				
		XmlObject localPageSequenceProperties = null;
		
    cursor.push();
    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-sequence-properties"))) {
      localPageSequenceProperties = cursor.getObject();
    }
    cursor.pop();
    
    if (localPageSequenceProperties == null) {
      localPageSequenceProperties = docPageSequenceProperties;
    }
		
    cursor.push();
		cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "body"));
		XWPFParagraph lastPara = handleBody(doc, cursor.getObject(), localPageSequenceProperties);
		cursor.pop();
		
    if (log.isDebugEnabled()) {
      // log.debug("handleSection(): Setting sectPr on last paragraph.");
    }
    CTPPr ppr = (lastPara.getCTP().isSetPPr() ? lastPara.getCTP().getPPr() : lastPara.getCTP().addNewPPr()); 
    CTSectPr sectPr = ppr.addNewSectPr();

    String sectionType = cursor.getAttributeText(DocxConstants.QNAME_TYPE_ATT);

    if (sectionType != null) {
      CTSectType type = sectPr.addNewType();
      type.setVal(STSectionMark.Enum.forString(sectionType));
    }

    setupPageSequence(doc, localPageSequenceProperties, sectPr);
    
    ppr.setSectPr(sectPr);        
		
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
    if (null != widthVal) {
      try {
        long width = Measurement.toTwips(widthVal, getDotsPerInch());
        pageSize.setW(BigInteger.valueOf(width));
      } catch (MeasurementException e) {
        log.warn("setPageSize(): Value \"" + widthVal + " for attribute \"width\" cannot be converted to a twips value");
      }
    }

    String heightVal = cursor.getAttributeText(DocxConstants.QNAME_HEIGHT_ATT);
    if (null != heightVal) {
      try {
        long height = Measurement.toTwips(heightVal, getDotsPerInch());
        pageSize.setH(BigInteger.valueOf(height));
      } catch (MeasurementException e) {
        log.warn("setPageSize(): Value \"" + heightVal + " for attribute \"height\" cannot be converted to a twips value");
      }
    }
  }

  /**
	 * Set up page sequence properties for the entire document, including page geometry, numbering, and headers and footers.
	 * @param doc Document to be constructed
	 * @param xml page-sequence-properties element
	 * @param sectPr Section properties to store the page sequence details on.
	 * @throws DocxGenerationException 
	 */
	private void setupPageSequence(XWPFDocument doc, XmlObject xml) throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();
		
		CTDocument1 document = doc.getDocument();
		CTBody body = (document.isSetBody() ? document.getBody() : document.addNewBody());
		CTSectPr sectPr = (body.isSetSectPr() ? body.getSectPr() : body.addNewSectPr());
		
		setPageNumberProperties(cursor, sectPr);
		cursor.push();
		if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "headers-and-footers"))) {
			constructHeadersAndFooters(doc, cursor.getObject());
		}
		cursor.pop();
    cursor.push();
    if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-size"))) {
      setPageSize(cursor, sectPr);
    }
    cursor.pop();
		
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
					  titlePg.setVal(STOnOff.TRUE);
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
            titlePg.setVal(STOnOff.TRUE);
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
	private void makeRun(XWPFParagraph para, XmlObject xml) throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();
		
		// String tagname = cursor.getName().getLocalPart(); // For debugging
		
		XWPFRun run = para.createRun();
		String styleName = cursor.getAttributeText(DocxConstants.QNAME_STYLE_ATT);
		String styleId = cursor.getAttributeText(DocxConstants.QNAME_STYLEID_ATT);
		
		if (null != styleName && null == styleId) {
			// Look up the style by name:
			XWPFStyle style = para.getDocument().getStyles().getStyleWithName(styleName);
			if (null != style) {
				styleId = style.getStyleId();
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
				run.setText(cursor.getTextValue());
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
				} else {
					log.error("makeRun(); Unexpected element {" + namespace + "}:" + name + ". Skipping.");
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
				      onOff.setVal(STOnOff.Enum.forString(attValue));
				      run.getCTR().getRPr().setOutline(onOff);
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
	 */
	private void makeFootnote(XWPFParagraph para, XmlObject xml) throws DocxGenerationException {
		
	  XmlCursor cursor = xml.newCursor();
	  
		String type = cursor.getAttributeText(DocxConstants.QNAME_TYPE_ATT);
		
		XWPFAbstractFootnoteEndnote note = null;
		if ("endnote".equals(type)) {
			note = para.getDocument().createEndnote();
		} else {
			note = para.getDocument().createFootnote();
		}
		
		// NOTE: The paragraph is not created with any initial paragraph.
		
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
		bookmark.setName(cursor.getAttributeText(DocxConstants.QNAME_NAME_ATT));
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
	 * Construct a hyperlink
	 * @param doc
	 * @param cursor
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
		
		CTHyperlink hyperlink = para.getCTP().addNewHyperlink();
		CTR run = hyperlink.addNewR();
		run.addNewT().setStringValue(cursor.getTextValue());
		
		// Set the appropriate target:
		
		if (href.startsWith("#")) {
			// Just a fragment ID, must be to a bookmark
			String bookmarkName = href.substring(1);
			hyperlink.setAnchor(bookmarkName);
		} else {
			// Create a relationship that targets the href and use the
			// relationship's ID on the hyperlink
			// It's not yet clear from the POI API how to create a new relationship for
			// use by an external hyperlink.
			// throw new NotImplementedException("Links to external resources not yet implemented.");
		}
		
		XWPFHyperlinkRun hyperlinkRun = new XWPFHyperlinkRun(hyperlink, run, para);
		para.addRun(hyperlinkRun);
		
	}

	/**
	 * Construct an image reference
	 * @param doc
	 * @param cursor
	 */
	private void makeImage(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
		cursor.push();
		
		String imgUrl = cursor.getAttributeText(DocxConstants.QNAME_SRC_ATT);
		if (null == imgUrl) {
			log.error("- [ERROR] No @src attribute for image.");
			return;
		}
		URL url;
		try {
			if (!imgUrl.matches("^\\w+:.+")) {
				String baseUrl = inFile.getParentFile().toURI().toURL().toExternalForm();
				imgUrl = baseUrl + imgUrl;
			}
			url = new URL(imgUrl);
		} catch (MalformedURLException e) {
			log.error("- [ERROR] " + e.getClass().getSimpleName() + " on img/@src value: " + e.getMessage());
			return;
		}
		File imgFile = null;
		try {
			imgFile = new File(url.toURI());
		} catch (URISyntaxException e) {
			// Should never get here.
		}
		
		String imgFilename = imgFile.getName();
		String imgExtension = FilenameUtils.getExtension(imgFilename).toLowerCase();
		int width = 200; // Default width in pixels
		int height = 200; // Default height in pixels
		
		int format = getImageFormat(imgExtension);
		
        if (format == 0) {
        	// FIXME: Might be more appropriate to throw an exception here.
            log.error("Unsupported picture: " + imgFilename +
                    ". Expected emf|wmf|pict|jpeg|jpg|png|dib|gif|tiff|eps|bmp|wpg");
            cursor.pop();
            return;
        }

		BufferedImage img = null;
		int intrinsicWidth = 0;
		int intrinsicHeight = 0;
		try {		
			// FIXME: Need to limit this to the formats Java2D can read.
		    img = ImageIO.read(imgFile);
		    intrinsicWidth = img.getWidth();
		    intrinsicHeight = img.getHeight();
		} catch (IOException e) {
			log.warn("" + e.getClass().getSimpleName() + " exception loading image file '" + imgFile +"': " +
                     e.getMessage());
		}		
		 
		String widthVal = cursor.getAttributeText(DocxConstants.QNAME_WIDTH_ATT);
		if (null != widthVal) {
			try {
				width = (int) Measurement.toPixels(widthVal, getDotsPerInch());
			} catch (MeasurementException e) {
				log.error(e.getClass().getSimpleName() + ": " + e.getMessage());
				log.error("Using default width value " + width);
				width = intrinsicWidth > 0 ? intrinsicWidth : width;
			}
		} else {
			width = intrinsicWidth > 0 ? intrinsicWidth : width;			
		}

		String heightVal = cursor.getAttributeText(DocxConstants.QNAME_HEIGHT_ATT);
		if (null != heightVal) {
			try {
				height = (int) Measurement.toPixels(heightVal, getDotsPerInch());
			} catch (MeasurementException e) {
				log.error(e.getClass().getSimpleName() + ": " + e.getMessage());
				log.error("Using default height value " + height);
				height = intrinsicHeight > 0 ? intrinsicHeight : height;
			}
		} else {
			height = intrinsicHeight > 0 ? intrinsicHeight : height;
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
			run.addPicture(new FileInputStream(imgFile), 
					       format, 
					       imgFilename, 
					       Units.toEMU(width), 
					       Units.toEMU(height));
		} catch (Exception e) {
			log.warn("" + e.getClass().getSimpleName() + " exception adding picture for reference '" + imgFile +"': " +
 		                       e.getMessage());
		}
		cursor.pop();
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
		// 		cursor.push();
		//		cursor.pop();
		
	}

	/**
	 * Construct an embedded object.
	 * @param doc
	 * @param cursor Cursor pointing to an <object> element.
	 */
	private void makeObject(XWPFDocument doc, XmlCursor cursor) throws DocxGenerationException {
		throw new NotImplementedException("Object handling not implemented");
		// 		cursor.push();
		//		cursor.pop();
		
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
		if (null != widthValue) {
		    table.setWidth(getMeasurementValue(widthValue));
		}
		
		setTableIndents(table, cursor);
		
    String styleName = cursor.getAttributeText(DocxConstants.QNAME_STYLE_ATT);
    String styleId = cursor.getAttributeText(DocxConstants.QNAME_STYLEID_ATT);
    
    if (null != styleName && null == styleId) {
      // Look up the style by name:
      XWPFStyle style = table.getBody().getXWPFDocument().getStyles().getStyleWithName(styleName);
      if (null != style) {
        styleId = style.getStyleId();
      } else {
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
				if (null != width) {
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
			
			try {
				String widthValue = cursor.getAttributeText(DocxConstants.QNAME_WIDTH_ATT);
				if (null != widthValue) {
					cell.setWidth(TableColumnDefinition.interpretWidthSpecification(widthValue, getDotsPerInch()));
				} else {
				  String width = null;
          width = colDef.getWidth();
				  if (colspan != null) {
				    try {
              long spanCount = Integer.parseInt(colspan);
              // Try to add up the widths of the spanned columns.
              // This is only possible if the values are all percents
              // or are all measurements. Since we don't the actual
              // width of the table itself necessarily, there's no way
              // to reliably convert percentages to explicit widths.
              List<String> spanWidths = new ArrayList<String>();
              for (int i = cellCtr; i < cellCtr + spanCount; i++) {
                spanWidths.add(colDefs.get(i).getSpecifiedWidth());
              }
              boolean allPercents = true;
              for (String cand : spanWidths) {
                allPercents = allPercents && cand.endsWith("%");
              }
              boolean allNumbers = true;
              for (String cand : spanWidths) {
                allNumbers = allNumbers && !cand.endsWith("%") && !cand.equals("auto");
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
                log.warn("Widths of spanned columns are neither all percents or all measurements, cannot calculate exact spanned width");
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
				if (cursor.toChild(DocxConstants.QNAME_P_ELEM)) {
					do {
						XWPFParagraph p = cell.addParagraph();
						makeParagraph(p, cursor);
						if (null != align) {
							ParagraphAlignment alignment = ParagraphAlignment.valueOf(align.toUpperCase());
							p.setAlignment(alignment);
						}
					} while(cursor.toNextSibling());
					// Cells always have at least one paragraph.
					cell.removeParagraph(0);
				}
			}
			cursor.pop();
			cellCtr++;
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
      }
      if (borderStyles.getTopBorder() != null) {
        CTBorder top = borders.addNewTop();
        top.setVal(borderStyles.getTopBorderEnum());
      }
      if (borderStyles.getLeftBorder() != null) {
        CTBorder left = borders.addNewLeft();
        left.setVal(borderStyles.getLeftBorderEnum());
      }
      if (borderStyles.getRightBorder() != null) {
        CTBorder right = borders.addNewRight();
        right.setVal(borderStyles.getRightBorderEnum());
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
		style.setCustomStyle(STOnOffImpl.X_1);
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
