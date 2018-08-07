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
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlCursor.TokenType;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalAlignRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.STOnOffImpl;

/**
 * Generates DOCX files from Simple Word Processing Markup Language XML.
 */
public class DocxGenerator {
	
	public static final Logger log = LogManager.getLogger();

	private static final String OO_WPML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	public static final String SIMPLE_WP_NS = "urn:ns:wordinator:simplewpml";
	
	private static final QName QNAME_INSTR_ATT = new QName(OO_WPML_NS, "instr");
	private static final QName QNAME_ALIGN_ATT = new QName("", "align");
	private static final QName QNAME_BOLD_ATT = new QName("", "bold");
	private static final QName QNAME_BOTTOM_ATT = new QName("", "bottom");
	private static final QName QNAME_CALCULATEDWIDTH_ATT = new QName("", "calculatedWidth");
	private static final QName QNAME_CAPS_ATT = new QName("", "caps");
	private static final QName QNAME_COLSEP_ATT = new QName("", "colsep");
	private static final QName QNAME_COLSPAN_ATT = new QName("", "colspan");
	private static final QName QNAME_COLWIDTH_ATT = new QName("", "colwidth");
	private static final QName QNAME_DOUBLE_STRIKETHROUGH_ATT = new QName("", "double-strikethrough");
	private static final QName QNAME_EMBOSS_ATT = new QName("", "emboss");
	private static final QName QNAME_EMPHASIS_MARK_ATT = new QName("", "emphasis-mark");
	private static final QName QNAME_EXPAND_COLLAPSE_ATT = new QName("", "expand-collapse");
	private static final QName QNAME_FONT_ATT = new QName("", "font");
	private static final QName QNAME_FOOTER_ATT = new QName("", "footer");
	private static final QName QNAME_FORMAT_ATT = new QName("", "format");
	private static final QName QNAME_FRAME_ATT = new QName("", "frame");
	private static final QName QNAME_GUTTER_ATT = new QName("", "gutter");
	private static final QName QNAME_HEADER_ATT = new QName("", "header");
	private static final QName QNAME_HEIGHT_ATT = new QName("", "height");
	private static final QName QNAME_HIGHLIGHT_ATT = new QName("", "highlight");
	private static final QName QNAME_HREF_ATT = new QName("", "href");
	private static final QName QNAME_ID_ATT = new QName("", "id");
	private static final QName QNAME_IMPRINT_ATT = new QName("", "imprint");
	private static final QName QNAME_ITALIC_ATT = new QName("", "italic");
	private static final QName QNAME_LEFT_ATT = new QName("", "left");
	private static final QName QNAME_NAME_ATT = new QName("", "name");
	private static final QName QNAME_OUTLINE_ATT = new QName("", "outline");
	private static final QName QNAME_OUTLINE_LEVEL_ATT = new QName("", "outline-level");
	private static final QName QNAME_PAGE_BREAK_BEFORE_ATT = new QName("", "page-break-before");
	private static final QName QNAME_POSITION_ATT = new QName("", "position");
	private static final QName QNAME_RIGHT_ATT = new QName("", "right");
	private static final QName QNAME_ROWSEP_ATT = new QName("", "rowsep");
	private static final QName QNAME_ROWSPAN_ATT = new QName("", "rowspan");
	private static final QName QNAME_SHADOW_ATT = new QName("", "shadow");
	private static final QName QNAME_SMALL_CAPS_ATT = new QName("", "small-caps");
	private static final QName QNAME_SRC_ATT = new QName("", "src");
	private static final QName QNAME_START_ATT = new QName("", "start");
	private static final QName QNAME_STRIKETHROUGH_ATT = new QName("", "strikethrough");
	private static final QName QNAME_STYLE_ATT = new QName("", "style");
	private static final QName QNAME_STYLEID_ATT = new QName("", "styleId");
	private static final QName QNAME_TAGNAME_ATT = new QName("", "tagName");
	private static final QName QNAME_TOP_ATT = new QName("", "top");
	private static final QName QNAME_TYPE_ATT = new QName("", "type");
	private static final QName QNAME_UNDERLINE_ATT = new QName("", "underline");
	private static final QName QNAME_UNDERLINE_COLOR_ATT = new QName("", "underline-color");
	private static final QName QNAME_VALIGN_ATT = new QName("", "valign");
	private static final QName QNAME_VANISH_ATT = new QName("", "vanish");
	private static final QName QNAME_VERTICAL_ALIGNMENT_ATT = new QName("", "vertical-alignment");
	private static final QName QNAME_WIDTH_ATT = new QName("", "width");
	private static final QName QNAME_XSLT_FORMAT_ATT = new QName("", "xslt-format");	 
	private static final QName QNAME_COLS_ELEM = new QName(SIMPLE_WP_NS, "cols");
	@SuppressWarnings("unused")
	private static final QName QNAME_COL_ELEM = new QName(SIMPLE_WP_NS, "col");
	private static final QName QNAME_THEAD_ELEM = new QName(SIMPLE_WP_NS, "thead");
	private static final QName QNAME_TBODY_ELEM = new QName(SIMPLE_WP_NS, "tbody");
	@SuppressWarnings("unused")
	private static final QName QNAME_TR_ELEM = new QName(SIMPLE_WP_NS, "tr");
	private static final QName QNAME_TD_ELEM = new QName(SIMPLE_WP_NS, "td");
	private static final QName QNAME_P_ELEM = new QName(SIMPLE_WP_NS, "p");
	@SuppressWarnings("unused")
	private static final QName QNAME_ROW_ELEM = new QName(SIMPLE_WP_NS, "row");
	private static final QName QNAME_VSPAN_ELEM = new QName(SIMPLE_WP_NS, "vspan");
	private File outFile;
	private int dotsPerInch = 72; /* DPI */
	private double dotsPerInchFactor = 1.0/dotsPerInch;
	// Map of source IDs to internal object IDs.
	private Map<String, BigInteger> bookmarkIdToIdMap = new HashMap<String, BigInteger>();
	private int idCtr = 0;
	private File templateFile;
	private File inFile;

	/**
	 * 
	 * @param inFile File representing input document.
	 * @param outFile File to write DOCX result to
	 * @param templateFile DOTX template to initialze result DOCX with (provides style definitions)
	 */
	public DocxGenerator(File inFile, File outFile, File templateFile) {
		this.inFile = inFile;
		this.outFile = outFile;		
		this.templateFile = templateFile;
	}

	/*
	 * Generate the DOCX file from the input Simple WP ML document. 
	 * @param xml The XmlObject that holds the Simple WP XML content
	 */
	public void generate(XmlObject xml) throws DocxGenerationException, XmlException, IOException {
		
		
		XWPFDocument doc = new XWPFDocument();
		
		setupStyles(doc);
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
		String tagName = cursor.getName().getLocalPart();
		cursor.push();
		if (cursor.toChild(new QName(SIMPLE_WP_NS, "page-sequence-properties"))) {
			setupPageSequence(doc, cursor.getObject());
		}
		cursor.pop();
		tagName = cursor.getName().getLocalPart();
		cursor.toChild(new QName(SIMPLE_WP_NS, "body"));
		tagName = cursor.getName().getLocalPart();
		handleBody(doc, cursor.getObject());
		
		
	}

	/**
	 * Process the elements in &lt;body&gt;
	 * @param doc Document to add paragraphs to.
	 * @param xml Body element
	 * @throws DocxGenerationException
	 */
	private void handleBody(XWPFDocument doc, XmlObject xml) throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();
		if (cursor.toFirstChild()) {
			do {
				String tagName = cursor.getName().getLocalPart();
				String namespace = cursor.getName().getNamespaceURI();
				if ("p".equals(tagName)) {
					XWPFParagraph p = doc.createParagraph();
					makeParagraph(p, cursor);
				} else if ("section".equals(tagName)) {
					handleSection(doc, cursor.getObject());
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
	}

	/**
	 * Handle a &lt;section&gt; element
	 * @param doc Document we're adding to
	 * @param xml &lt;section&gt; element
	 */
	@SuppressWarnings("unused")
	private void handleSection(XWPFDocument doc, XmlObject xml) throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();
		
		log.warn("Section-level headers and footers and page numbering not yet implemented.");
		// FIXME: The section-specific properties go in the first paragraph of the section.
		cursor.push();
		if (false && cursor.toChild(new QName(SIMPLE_WP_NS, "page-sequence-properties"))) {
			setupPageSequence(doc, cursor.getObject());
		}
		cursor.pop();
		
		cursor.toChild(new QName(SIMPLE_WP_NS, "body"));
		handleBody(doc, cursor.getObject());
		
	}

	/**
	 * Set up page sequence properties for the entire document, including page geometry, numbering, and headers and footers.
	 * @param doc Document to be constructed
	 * @param xml page-sequence-properties element
	 * @throws DocxGenerationException 
	 */
	private void setupPageSequence(XWPFDocument doc, XmlObject xml) throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();
		cursor.push();
		if (cursor.toChild(new QName(SIMPLE_WP_NS, "page-number-properties"))) {
			String format = cursor.getAttributeText(QNAME_FORMAT_ATT);
			if (null != format) {
				// FIXME: Not sure how to set this up with the POI API yet.
			}
		}
		cursor.pop();
		cursor.push();
		if (cursor.toChild(new QName(SIMPLE_WP_NS, "headers-and-footers"))) {
			constructHeadersAndFooters(doc, cursor.getObject());
		}
		cursor.pop();
		
	}

	/**
	 * Construct headers and footers for the document.
	 * @param doc Document to add headers and footers to.
	 * @param xml headers-and-footers element
	 * @throws DocxGenerationException 
	 */
	private void constructHeadersAndFooters(XWPFDocument doc, XmlObject xml) throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();
		
		boolean haveOddHeader = false;
		boolean haveEvenHeader = false;
		boolean haveOddFooter = false;
		boolean haveEvenFooter = false;
		
		if (cursor.toFirstChild()) {
			do {
				String tagName = cursor.getName().getLocalPart();
				String namespace = cursor.getName().getNamespaceURI();
				if ("header".equals(tagName)) {
					HeaderFooterType type = getHeaderFooterType(cursor);
					if (type == HeaderFooterType.DEFAULT) {
						haveOddHeader = true;
					}
					if (type == HeaderFooterType.EVEN) {
						haveEvenHeader = true;
					}
					XWPFHeader header = doc.createHeader(type);
					makeHeaderFooter(header, cursor.getObject());
				} else if ("footer".equals(tagName)) {
					HeaderFooterType type = getHeaderFooterType(cursor);
					if (type == HeaderFooterType.DEFAULT) {
						haveOddFooter = true;
					}
					if (type == HeaderFooterType.EVEN) {
						haveEvenFooter = true;
					}
					XWPFFooter footer = doc.createFooter(type);
					makeHeaderFooter(footer, cursor.getObject());
				} else {
					log.warn("Unexpected element {" + namespace + "}:" + tagName + " in <headers-and-footers>. Ignored.");
				}
			} while(cursor.toNextSibling());
		}
		
		if ((haveOddHeader || haveOddFooter) && 
			(haveEvenHeader || haveEvenFooter)) {
			doc.setEvenAndOddHeadings(true);
		}
		
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
		String typeName = cursor.getAttributeText(QNAME_TYPE_ATT);
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
	private XWPFParagraph makeParagraph(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
		
		cursor.push();
		String styleName = cursor.getAttributeText(QNAME_STYLE_ATT);
		String styleId = cursor.getAttributeText(QNAME_STYLEID_ATT);
		
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
					makeFootnote(para, cursor);
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
		String styleName = cursor.getAttributeText(QNAME_STYLE_ATT);
		String styleId = cursor.getAttributeText(QNAME_STYLEID_ATT);
		
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
			CTRPr pr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
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
	private void makeFootnote(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
		
		String type = cursor.getAttributeText(QNAME_TYPE_ATT);
		
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
		
		String typeValue = cursor.getAttributeText(QNAME_TYPE_ATT);
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
		bookmark.setName(cursor.getAttributeText(QNAME_NAME_ATT));
		BigInteger id = nextId();
		bookmark.setId(id);
		this.bookmarkIdToIdMap.put(cursor.getAttributeText(QNAME_ID_ATT), id);
	}

	/**
	 * Construct a bookmark end
	 * @param doc
	 * @param cursor
	 * @throws DocxGenerationException 
	 */
	private void makeBookmarkEnd(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
		CTMarkupRange bookmark = para.getCTP().addNewBookmarkEnd();
		String sourceID = cursor.getAttributeText(QNAME_ID_ATT);
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
		
		String href = cursor.getAttributeText(QNAME_HREF_ATT);
		
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
		
		String imgUrl = cursor.getAttributeText(QNAME_SRC_ATT);
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
		 
		String widthVal = cursor.getAttributeText(QNAME_WIDTH_ATT);
		if (null != widthVal) {
			try {
				width = (int) Measurement.toPixels(widthVal, dotsPerInch);
			} catch (MeasurementException e) {
				log.error(e.getClass().getSimpleName() + ": " + e.getMessage());
				log.error("Using default width value " + width);
				width = intrinsicWidth > 0 ? intrinsicWidth : width;
			}
		} else {
			width = intrinsicWidth > 0 ? intrinsicWidth : width;			
		}

		String heightVal = cursor.getAttributeText(QNAME_HEIGHT_ATT);
		if (null != heightVal) {
			try {
				height = (int) Measurement.toPixels(heightVal, dotsPerInch);
			} catch (MeasurementException e) {
				log.error(e.getClass().getSimpleName() + ": " + e.getMessage());
				log.error("Using default height value " + height);
				height = intrinsicHeight > 0 ? intrinsicHeight : height;
			}
		} else {
			height = intrinsicHeight > 0 ? intrinsicHeight : height;
		}

	    double widthInches = width * dotsPerInchFactor;
	    double heightInches = height * dotsPerInchFactor;
	    width = (int) (widthInches * dotsPerInch);
	    height = (int) (heightInches * dotsPerInch);
		
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
		// Set the column widths. In the DOCX markup this is
		// done in the "grid" (<w:tblGrid>)
		XmlCursor cursor = xml.newCursor();
		
		CTTblGrid grid = table.getCTTbl().getTblGrid();
		if (grid == null) {
			// Create a new grid
			grid = table.getCTTbl().addNewTblGrid();
		}
		List<BigInteger> colWidths = new ArrayList<BigInteger>();
		cursor.toChild(QNAME_COLS_ELEM);
		if (cursor.toFirstChild()) {
			do {
				// The grid is constructed with no columns, so we can just
				// add columns for each cols element.
				String width = cursor.getAttributeText(QNAME_COLWIDTH_ATT);
				if (null != width) {
					try {
						// Column widths are in twips (1/20th of a point), not EMUs
						int twips = Measurement.toTwips(width, dotsPerInch);
						BigInteger colWidth = new BigInteger(Integer.toString(twips));
						colWidths.add(colWidth);
						CTTblGridCol gridCol = grid.addNewGridCol();	
						gridCol.setW(colWidth);
					} catch (MeasurementException e) {
						log.warn("makeTable(): " + e.getClass().getSimpleName() + ": " + e.getMessage());
					}
				}
			} while (cursor.toNextSibling());
		}
		// populate the rows and cells.
		cursor = xml.newCursor();
		
		// Header rows:
		cursor.push();
		if (cursor.toChild(QNAME_THEAD_ELEM)) {
			if (cursor.toFirstChild()) {
				RowSpanManager rowSpanManager = new RowSpanManager();
				do {
					// Process the rows
					XWPFTableRow row = makeTableRow(table, cursor.getObject(), colWidths, rowSpanManager);
					row.setRepeatHeader(true);
				} while(cursor.toNextSibling());
			}
		}
		
		// Body rows:
		
		cursor = xml.newCursor();
		if (cursor.toChild(QNAME_TBODY_ELEM)) {
			if (cursor.toFirstChild()) {
				RowSpanManager rowSpanManager = new RowSpanManager();
				do {
					// Process the rows
					XWPFTableRow row = makeTableRow(table, cursor.getObject(), colWidths, rowSpanManager);
					// Adjust row as needed.
					row.getCtRow(); // For setting low-level properties.
				} while(cursor.toNextSibling());
			}
		}
		table.removeRow(0); // Remove the first row that's always added automatically (FIXME: This may not be needed any more)
	}
	
	/**
	 * Construct a table row
	 * @param table The table to add the row to
	 * @param xml The <row> element to add to the table
	 * @param colWidths List of columns widths in column order.
	 * @param rowSpanManager Manages setting vertical spanning across multiple rows.
	 * @return Constructed row object
	 * @throws DocxGenerationException 
	 */
	private XWPFTableRow makeTableRow(
			XWPFTable table, 
			XmlObject xml, 
			List<BigInteger> colWidths, 
			RowSpanManager rowSpanManager) 
					throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();
		XWPFTableRow row = table.createRow();
		
		// FIXME: Handle attributes on rows (rowsep, colsep, etc.)
		
		cursor.push();
		cursor.toChild(QNAME_TD_ELEM);
		int cellCtr = 0;
		
		do {
			// Rows always have at least one cell
			// FIXME: At some point the POI API will remove the automatic creation
			// of the first cell in a row.
			XWPFTableCell cell = cellCtr == 0 ? row.getCell(0) : row.addNewTableCell();
			
			CTTcPr ctTcPr = cell.getCTTc().addNewTcPr();
			String align = cursor.getAttributeText(QNAME_ALIGN_ATT);
			String valign = cursor.getAttributeText(QNAME_VALIGN_ATT);
			String colspan = cursor.getAttributeText(QNAME_COLSPAN_ATT);
			String rowspan = cursor.getAttributeText(QNAME_ROWSPAN_ATT);
			
			try {
				ctTcPr.addNewTcW().setW(colWidths.get(cellCtr));
			} catch (Exception e) {
				// There might not be column widths defined for the table,
				// in which case just silently ignore this.
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
			
			cursor.push();
			// The first cell of a span will already have a vertical span set for it.
			if (rowspan == null && cursor.toChild(QNAME_VSPAN_ELEM)) {
				int spansRemaining = rowSpanManager.includeCell(cellCtr);
				if (spansRemaining < 0) {
					log.warn("Found <vspan> when there should not have been one. Ignored.");
				} else {
					ctTcPr.setVMerge(CTVMerge.Factory.newInstance());
				}
			} else {
				if (cursor.toChild(QNAME_P_ELEM)) {
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

	private void setupStyles(XWPFDocument doc) throws DocxGenerationException {
		// Load template. For now this is hard coded but will need to be
		// parameterized
				
		// Copy the template's styles to result document:
				
		try {
			XWPFDocument templateDoc = new XWPFDocument(new FileInputStream(templateFile));
			XWPFStyles newStyles = doc.createStyles();
			newStyles.setStyles(templateDoc.getStyle());
			templateDoc.close();
		} catch (FileNotFoundException e) {
			throw new DocxGenerationException("setupStyles(): Expected DOCX template for styles not found: " + templateFile.getAbsolutePath(), e);
		} catch (IOException e) {
			new DocxGenerationException(e.getClass().getSimpleName() + " loading template DOCX file: " + e.getMessage(), e);
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
