/**
 *
 */
package org.wordinator.xml2docx.generator;

import java.util.ArrayList;
import java.util.Hashtable;
import java.util.List;

import javax.xml.namespace.QName;

/**
 * Constants for names and namespaces and such.
 *
 */
@SuppressWarnings("unused")
public final class DocxConstants {

  // Namespace names:
	public static final String OO_WPML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	public static final String SIMPLE_WP_NS = "urn:ns:wordinator:simplewpml";

	// Attributes:
	public static final QName QNAME_ALIGN_ATT = new QName("", "align");
  public static final QName QNAME_ARG_A_ATT = new QName("", "arg-a");
  public static final QName QNAME_ARG_B_ATT = new QName("", "arg-b");
  public static final QName QNAME_ARG_C_ATT = new QName("", "arg-c");
  public static final QName QNAME_ARG_D_ATT = new QName("", "arg-d");
  public static final QName QNAME_ARG_F_ATT = new QName("", "arg-f");
  public static final QName QNAME_ARG_H_ATT = new QName("", "arg-h");
  public static final QName QNAME_ARG_L_ATT = new QName("", "arg-l");
  public static final QName QNAME_ARG_N_ATT = new QName("", "arg-n");
  public static final QName QNAME_ARG_O_ATT = new QName("", "arg-o");
  public static final QName QNAME_ARG_P_ATT = new QName("", "arg-p");
  public static final QName QNAME_ARG_S_ATT = new QName("", "arg-s");
  public static final QName QNAME_ARG_T_ATT = new QName("", "arg-t");
  public static final QName QNAME_ARG_U_ATT = new QName("", "arg-u");
  public static final QName QNAME_ARG_W_ATT = new QName("", "arg-w");
  public static final QName QNAME_ARG_X_ATT = new QName("", "arg-x");
  public static final QName QNAME_ARG_Z_ATT = new QName("", "arg-z");
	public static final QName QNAME_BOLD_ATT = new QName("", "bold");
  public static final QName QNAME_BORDER_COLOR_ATT = new QName("", "bordercolor");
  public static final QName QNAME_BORDER_COLOR_TOP_ATT = new QName("", "bordercolortop");
  public static final QName QNAME_BORDER_COLOR_LEFT_ATT = new QName("", "bordercolorleft");
  public static final QName QNAME_BORDER_COLOR_BOTTOM_ATT = new QName("", "bordercolorbottom");
  public static final QName QNAME_BORDER_COLOR_RIGHT_ATT = new QName("", "bordercolorright");
  public static final QName QNAME_BORDER_STYLE_ATT = new QName("", "borderstyle");
  public static final QName QNAME_BORDER_STYLE_BOTTOM_ATT = new QName("", "borderstylebottom");
  public static final QName QNAME_BORDER_STYLE_LEFT_ATT = new QName("", "borderstyleleft");
  public static final QName QNAME_BORDER_STYLE_INSIDE_ATT = new QName("", "borderstyleinside");
  public static final QName QNAME_BORDER_STYLE_OUTSIDE_ATT = new QName("", "borderstyleoutside");
  public static final QName QNAME_BORDER_STYLE_RIGHT_ATT = new QName("", "borderstyleright");
  public static final QName QNAME_BORDER_STYLE_TOP_ATT = new QName("", "borderstyletop");
	public static final QName QNAME_BOTTOM_ATT = new QName("", "bottom");
	public static final QName QNAME_CALCULATEDWIDTH_ATT = new QName("", "calculatedWidth");
  public static final QName QNAME_CALLOUT_ATT = new QName("", "callout");
	public static final QName QNAME_CAPS_ATT = new QName("", "caps");
  public static final QName QNAME_CHAPTER_SEPARATOR_ATT = new QName("", "chapter-separator");
  public static final QName QNAME_CHAPTER_STYLE_ATT = new QName("", "chapter-style");
  public static final QName QNAME_CODE_ATT = new QName(SIMPLE_WP_NS, "code");
  public static final QName QNAME_COLOR_ATT = new QName(OO_WPML_NS, "color");
	public static final QName QNAME_COLSEP_ATT = new QName("", "colsep");
	public static final QName QNAME_COLSPAN_ATT = new QName("", "colspan");
	public static final QName QNAME_COLWIDTH_ATT = new QName("", "colwidth");
  public static final QName QNAME_CUSTOMMARKFOLLOWS_ATT = new QName(OO_WPML_NS, "customMarkFollows");
	public static final QName QNAME_DOUBLE_STRIKETHROUGH_ATT = new QName("", "double-strikethrough");
	public static final QName QNAME_EMBOSS_ATT = new QName("", "emboss");
	public static final QName QNAME_EMPHASIS_MARK_ATT = new QName("", "emphasis-mark");
	public static final QName QNAME_EXPAND_COLLAPSE_ATT = new QName("", "expand-collapse");
	public static final QName QNAME_FONT_ATT = new QName("", "font");
	public static final QName QNAME_FOOTER_ATT = new QName("", "footer");
	public static final QName QNAME_FORMAT_ATT = new QName("", "format");
	public static final QName QNAME_FRAME_ATT = new QName("", "frame");
	public static final QName QNAME_FRAMESTYLE_ATT = new QName("", "framestyle");
  public static final QName QNAME_FRAMESTYLE_BOTTOM_ATT = new QName("", "framestyleBottom");
  public static final QName QNAME_FRAMESTYLE_LEFT_ATT = new QName("", "framestyleLeft");
  public static final QName QNAME_FRAMESTYLE_RIGHT_ATT = new QName("", "framestyleRight");
  public static final QName QNAME_FRAMESTYLE_TOP_ATT = new QName("", "framestyleTop");
	public static final QName QNAME_GUTTER_ATT = new QName("", "gutter");
	public static final QName QNAME_HEADER_ATT = new QName("", "header");
	public static final QName QNAME_HEIGHT_ATT = new QName("", "height");
	public static final QName QNAME_HIGHLIGHT_ATT = new QName("", "highlight");
	public static final QName QNAME_HREF_ATT = new QName("", "href");
	public static final QName QNAME_ID_ATT = new QName("", "id");
	public static final QName QNAME_IMPRINT_ATT = new QName("", "imprint");
  public static final QName QNAME_INSIDEINDENT_ATT = new QName("", "insideindent");
  public static final QName QNAME_INSTR_ATT = new QName(OO_WPML_NS, "instr");
	public static final QName QNAME_ITALIC_ATT = new QName("", "italic");
	public static final QName QNAME_KEEPLINES_ATT = new QName("", "keeplines");
	public static final QName QNAME_KEEPNEXT_ATT = new QName("", "keepnext");
  public static final QName QNAME_LAYOUT_ATT = new QName("", "layout");
	public static final QName QNAME_LEFT_ATT = new QName("", "left");
  public static final QName QNAME_LEFTINDENT_ATT = new QName("", "leftindent");
	public static final QName QNAME_NAME_ATT = new QName("", "name");
  public static final QName QNAME_ORIENT_ATT = new QName("", "orient");
	public static final QName QNAME_OUTLINE_ATT = new QName("", "outline");
	public static final QName QNAME_OUTLINE_LEVEL_ATT = new QName("", "outline-level");
  public static final QName QNAME_OUTSIDEINDENT_ATT = new QName("", "outsideindent");
	public static final QName QNAME_PAGE_BREAK_BEFORE_ATT = new QName("", "page-break-before");
	public static final QName QNAME_POSITION_ATT = new QName("", "position");
  public static final QName QNAME_REFERENCE_CALLOUT_ATT = new QName("", "reference-callout");
	public static final QName QNAME_RIGHT_ATT = new QName("", "right");
  public static final QName QNAME_RIGHTINDENT_ATT = new QName("", "rightindent");
	public static final QName QNAME_ROWSEP_ATT = new QName("", "rowsep");
	public static final QName QNAME_ROWSPAN_ATT = new QName("", "rowspan");
  public static final QName QNAME_SHADE_ATT = new QName("", "shade");
	public static final QName QNAME_SHADOW_ATT = new QName("", "shadow");
	public static final QName QNAME_SMALL_CAPS_ATT = new QName("", "small-caps");
	public static final QName QNAME_SRC_ATT = new QName("", "src");
	public static final QName QNAME_START_ATT = new QName("", "start");
	public static final QName QNAME_STRIKETHROUGH_ATT = new QName("", "strikethrough");
	public static final QName QNAME_STYLE_ATT = new QName("", "style");
	public static final QName QNAME_STYLEID_ATT = new QName("", "styleId");
	public static final QName QNAME_TAGNAME_ATT = new QName("", "tagName");
	public static final QName QNAME_TOP_ATT = new QName("", "top");
	public static final QName QNAME_TYPE_ATT = new QName("", "type");
	public static final QName QNAME_UNDERLINE_ATT = new QName("", "underline");
	public static final QName QNAME_UNDERLINE_COLOR_ATT = new QName("", "underline-color");
	public static final QName QNAME_VALIGN_ATT = new QName("", "valign");
	public static final QName QNAME_VANISH_ATT = new QName("", "vanish");
	public static final QName QNAME_VERTICAL_ALIGNMENT_ATT = new QName("", "vertical-alignment");
	public static final QName QNAME_WIDTH_ATT = new QName("", "width");
	public static final QName QNAME_XSLT_FORMAT_ATT = new QName("", "xslt-format");

	// Elements:
	public static final QName QNAME_COLS_ELEM = new QName(SIMPLE_WP_NS, "cols");
	public static final QName QNAME_COL_ELEM = new QName(SIMPLE_WP_NS, "col");
	public static final QName QNAME_CORE_PROPERTIES_ELEM = new QName(SIMPLE_WP_NS, "core-properties");
	public static final QName QNAME_CUSTOM_PROPERTIES_ELEM = new QName(SIMPLE_WP_NS, "custom-properties");
	public static final QName QNAME_DOCUMENT_PROPERTIES_ELEM = new QName(SIMPLE_WP_NS, "document-properties");
	public static final QName QNAME_EXTENDED_PROPERTIES_ELEM = new QName(SIMPLE_WP_NS, "extended-properties");
	public static final QName QNAME_FIELDRESULTS_ELEM = new QName(SIMPLE_WP_NS, "fieldResults");
	public static final QName QNAME_FOOTNOTEREF_ELEM = new QName(OO_WPML_NS, "footnoteRef");
	public static final QName QNAME_FOOTNOTEREFEREMCE_ELEM = new QName(OO_WPML_NS, "footnoteReference");
	public static final QName QNAME_INSTRUCTIONTEXT_ELEM = new QName(SIMPLE_WP_NS, "instructionText");
	public static final QName QNAME_P_ELEM = new QName(SIMPLE_WP_NS, "p");
	public static final QName QNAME_W_P_ELEM = new QName(OO_WPML_NS, "p");
	public static final QName QNAME_R_ELEM = new QName(OO_WPML_NS, "r");
	public static final QName QNAME_ROW_ELEM = new QName(SIMPLE_WP_NS, "row");
	public static final QName QNAME_T_ELEM = new QName(OO_WPML_NS, "t"); // w:t -- text element
	public static final QName QNAME_TABLE_ELEM = new QName(SIMPLE_WP_NS, "table"); // w:table -- table element
	public static final QName QNAME_THEAD_ELEM = new QName(SIMPLE_WP_NS, "thead");
	public static final QName QNAME_TBODY_ELEM = new QName(SIMPLE_WP_NS, "tbody");
	public static final QName QNAME_TR_ELEM = new QName(SIMPLE_WP_NS, "tr");
	public static final QName QNAME_TD_ELEM = new QName(SIMPLE_WP_NS, "td");
	public static final QName QNAME_VSPAN_ELEM = new QName(SIMPLE_WP_NS, "vspan");
	
	public static final String PROPERTY_VALUE_CONTINUOUS = "continuous";
  public static final String PROPERTY_PAGEBREAK = "pagebreak";
  public static final QName QNAME_TCPR_ELEM = new QName(OO_WPML_NS, "tcPr");
  public static final QName QNAME_TCBORDERS_ELEM = new QName(OO_WPML_NS, "tcBorders");
  public static final QName QNAME_PGSZ_ELEM = new QName(OO_WPML_NS, "pgSz");
  public static final QName QNAME_TOP_ELEM = new QName(OO_WPML_NS, "top");
  public static final QName QNAME_VAL_ATT = new QName(OO_WPML_NS, "val");
  public static final QName QNAME_OOXML_LEFT_ATT = new QName(OO_WPML_NS, "left");
  public static final QName QNAME_OOXML_RIGHT_ATT = new QName(OO_WPML_NS, "right");
  public static final QName QNAME_OOXML_TOP_ATT = new QName(OO_WPML_NS, "top");
  public static final QName QNAME_OOXML_BOTTOM_ATT = new QName(OO_WPML_NS, "bottom");
  public static final QName QNAME_OOXML_ORIENT_ATT = new QName(OO_WPML_NS, "orient");
  public static final QName QNAME_OOXML_H_ATT = new QName(OO_WPML_NS, "h");
  public static final QName QNAME_OOXML_W_ATT = new QName(OO_WPML_NS, "w");
  public static final QName QNAME_LEFT_ELEM = new QName(OO_WPML_NS, "left");
  public static final QName QNAME_RIGHT_ELEM = new QName(OO_WPML_NS, "right");
  public static final QName QNAME_BOTTOM_ELEM = new QName(OO_WPML_NS, "bottom");
  public static final QName QNAME_TBLPR_ELEM = new QName(OO_WPML_NS, "tblPr");
  public static final QName QNAME_TBLLAYOUT_ELEM = new QName(OO_WPML_NS, "tblLayout");
  public static final QName QNAME_WTYPE_ATT = new QName(OO_WPML_NS, "type");


}
