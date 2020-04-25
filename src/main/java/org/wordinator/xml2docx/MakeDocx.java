package org.wordinator.xml2docx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import javax.xml.transform.Source;
import javax.xml.transform.stream.StreamSource;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.filefilter.SuffixFileFilter;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlObject;
import org.wordinator.xml2docx.generator.DocxGeneratingOutputUriResolver;
import org.wordinator.xml2docx.generator.DocxGenerator;
import org.wordinator.xml2docx.saxon.Log4jSaxonLogger;
import org.wordinator.xml2docx.saxon.LoggingMessageListener;

import net.sf.saxon.lib.Feature;
import net.sf.saxon.lib.StandardErrorListener;
import net.sf.saxon.s9api.MessageListener2;
import net.sf.saxon.s9api.Processor;
import net.sf.saxon.s9api.QName;
import net.sf.saxon.s9api.XdmNode;
import net.sf.saxon.s9api.XdmValue;
import net.sf.saxon.s9api.Xslt30Transformer;
import net.sf.saxon.s9api.XsltCompiler;
import net.sf.saxon.s9api.XsltExecutable;
import net.sf.saxon.trans.XPathException;
import net.sf.saxon.trans.XmlCatalogResolver;

/**
 * Command-line application to generate DOCX files from
 * 	
 * <p>You can use this directly as the main file run from the command line
 * or as a helper class to build your own command-line handler or integrated
 * DOCX generator.
 *
 */
public class MakeDocx 
{
	
	private static final String APACHE_RESOLVER_CLASS = "org.apache.xml.resolver.CatalogManager";

  public static final String OPTION_CHAR_CHUNKLEVEL = "c";

  public static final String OPTION_CHAR_CATALOG = "k";

  public static final String OPTION_CHAR_TRANSFORMPATH = "x";

  public static final String OPTION_CHAR_TEMPLATEPATH = "t";

  public static final String OPTION_CHAR_OUTPUTPATH = "o";

  public static final String OPTION_CHAR_INPUTPATH = "i";

  public static final Logger log = LogManager.getLogger(MakeDocx.class.getSimpleName());
			
	public static final String XSLT_PARAM_CHUNKLEVEL = "chunklevel";

	public static void main( String[] args ) throws ParseException
    {
	    boolean GOOD_OPTIONS = false;
	    Options options = null;
	    try {
	      options = buildOptions();
	      GOOD_OPTIONS = true;
	    } catch (Exception e) {
	      // 
	    }
    	
    	try {
    	  handleCommandLine(options, args, log);
    	} catch (ParseException e) {
    	  GOOD_OPTIONS = false;
    	} catch (Exception e) {
    	  log.error(e.getClass().getSimpleName() + ": " + e.getMessage());
    	  System.exit(1);
      }

    	if (!GOOD_OPTIONS) {
        HelpFormatter formatter = new HelpFormatter();
        formatter.printHelp( "wordinator", options, true );    	  
    	}
    }

	/**
	 * Does the actual command line processing. You can call this from your own
	 * command line processor if you need additional command-line options, for example,
	 * to set additional XSLT parameters.
	 * @param options Command-line options
	 * @param args Command-line arguments
	 * @param log Logger to log messages to.
	 * @throws ParseException Thrown if there is problem parsing the input
	 */
	public static void handleCommandLine(
			Options options,
			String[] args,
			Logger log) throws Exception {
    	CommandLineParser parser = new DefaultParser();
    	CommandLine cmd = parser.parse( options, args);
    	
    	Map<String, String> xsltParameters = new HashMap<String, String>();
		  String inDocPath = cmd.getOptionValue(OPTION_CHAR_INPUTPATH).trim();
    	String docxPath = cmd.getOptionValue(OPTION_CHAR_OUTPUTPATH).trim();
    	String templatePath = cmd.getOptionValue(OPTION_CHAR_TEMPLATEPATH).trim();
    	String transformPath = cmd.getOptionValue(OPTION_CHAR_TRANSFORMPATH);
    	if (transformPath != null) {
    	  transformPath = transformPath.trim();
    	}
      String catalog = cmd.getOptionValue(OPTION_CHAR_CATALOG);
      if (catalog != null) {
        catalog = catalog.trim();
      }
    	String chunkLevel = cmd.getOptionValue(OPTION_CHAR_CHUNKLEVEL);
    	if (chunkLevel != null) {
    	  chunkLevel = chunkLevel.trim();
    	}
    	
    	chunkLevel = chunkLevel == null ? "root" : chunkLevel;
    	
    	log.info("Input document or directory='" + inDocPath + "'");
    	log.info("Output directory           ='" + docxPath + "'");
    	log.info("DOTX template              ='" + templatePath + "'");
    	log.info("XSLT template              =" + (transformPath == null ? "Not specified" : "'" + transformPath + "'"));
      log.info("Catalog                    =" + (catalog == null ? "Not specified" : "'" + catalog + "'"));
    	log.info("Chunk level                ='" + chunkLevel + "'");
    	
    	// Check that the input file exists.
    	// For now, always overwriting the DOCX file without confirmation.
    	
    	File inFile = new File(inDocPath);
    	if (!inFile.exists()) {
    		throw new RuntimeException("Input file '" + inFile.getAbsolutePath() + "' not found. Cannot continue."); 
    	}
    	File templateFile = new File(templatePath);
    	if (!templateFile.exists()) {
    	  throw new RuntimeException("Template file '" + templateFile.getAbsolutePath() + "' not found. Cannot continue."); 
    	}
    	
		XWPFDocument templateDoc = null;
		try {
			templateDoc = new XWPFDocument(new FileInputStream(templateFile));
		} catch (Exception e) {
		  throw new RuntimeException(e.getClass().getSimpleName() +  " loading template DOCX file \"" + templateFile.getAbsolutePath() + "\"");
		}
    	
    	File outFile = new File(docxPath);
    	
    	File outDir = outFile; // Normal case: specify output directory
    	if (outFile.getName().endsWith(".docx")) {
    		outDir = outFile.getParentFile();
    	}
    	
    	if (!outDir.exists()) {
    		log.info("Making output directory '" + outDir.getAbsolutePath() + "'...");
    		if (!outDir.mkdirs()) {
    		  try {
            templateDoc.close();
          } catch (IOException e) {
            // Don't care about this should it ever happen.
          }
    		  throw new RuntimeException("Failed to create output directory '" + outDir.getAbsolutePath() + "'. Cannot continue");
    		}
    	}
    	
    	File transformFile = null;
    	if (null != transformPath) {
    		transformFile = new File(transformPath);
        	if (!transformFile.exists()) {
        	  try {
              templateDoc.close();
            } catch (IOException e) {
              // Don't care about this should it ever happen.
            }
        	  throw new RuntimeException("XSLT transform file '" + transformFile.getAbsolutePath() + "' not found. Cannot continue."); 
        	}
        	if (!xsltParameters.containsKey(XSLT_PARAM_CHUNKLEVEL)) {
        		xsltParameters.put(XSLT_PARAM_CHUNKLEVEL, chunkLevel);
        	}
    	}
    	
    	try {
    		if (inFile.isDirectory()) {
    			// Assume directory contains *.swpx files 
    			handleDirectory(inFile, outDir, templateDoc, log);
    		} else { 
    			if (inFile.getName().endsWith(".swpx")) {
	    			handleSingleSwpxDoc(inFile, outFile, templateDoc, log);
	    		} else {
	    			transformXml(inFile, outDir, templateDoc, transformFile, catalog, xsltParameters, log);
	    		}
    		}
    	} catch (Exception e) {
    	  throw new RuntimeException(e.getClass().getSimpleName() + ": " + e.getMessage(), e);
    	} finally {
			try {
				templateDoc.close();
			} catch (IOException e) {
				// Don't care about this should it ever happen.
			}
		}
    	
	}

	/**
	 * Process an XML document to a set of DOCX files
	 * @param docFile the root XML document to process
	 * @param outDir Directory to put the DOCX files in
	 * @param templateDoc Template DOCX document
	 * @param transformFile The file containing the XSLT transform for generating SWPX documents
	 * @param catalog List of catalog files (as for Saxon -catalog option). Maybe null.
	 * @param xsltParameters Map of parameter names to values to be passed to the XSLT transform.
	 * @param log Log to write messages to.
	 * @throws Exception Any kind of error
	 */
	public static void transformXml(
			File docFile, 
			File outDir, 
			XWPFDocument templateDoc, 
			File transformFile, 
			String catalog, 
			Map<String, String> xsltParameters, 
			Logger log) throws Exception {
		// Apply transform to book file to generate Simple WP XML documents
		
		if (transformFile == null) {
			throw new RuntimeException("-x (transform) parameter not specified. If the input is a _Book.xml file, you must specify the -x parameter");
		}
		
		StandardErrorListener errorListener = new StandardErrorListener();
		net.sf.saxon.lib.Logger saxonLogger = new Log4jSaxonLogger(log);
		errorListener.setLogger(saxonLogger);		
		
		Processor processor = new Processor(false);
		DocxGeneratingOutputUriResolver outputResolver = new DocxGeneratingOutputUriResolver(outDir, templateDoc, log);
		// Saxon 9.9+ version:
		processor.setConfigurationProperty(Feature.OUTPUT_URI_RESOLVER, outputResolver);
    // processor.setConfigurationProperty(FeatureKeys.OUTPUT_URI_RESOLVER, outputResolver);
		
    if (catalog != null) {
      // Adapted from Saxon CommandLineOptions.java:
      try {
        Class<?> klass = processor.getClass().getClassLoader().loadClass(APACHE_RESOLVER_CLASS);
        if (klass == null) {
          throw new RuntimeException("-k/-catalog option specified but failed to load class " + APACHE_RESOLVER_CLASS);
        }
        XmlCatalogResolver.setCatalog(catalog, processor.getUnderlyingConfiguration(), false);
      } catch (XPathException err) {
          throw new XPathException("Failed to load Apache catalog resolver library", err);
      }
    }
    

		
		
		// FIXME: Set up proper logger. See 
		// https://www.saxonica.com/html/documentation/using-xsl/embedding/s9api-transformation.html
		XsltCompiler compiler = processor.newXsltCompiler();
		
		InputStream inStream = new FileInputStream(transformFile);
		Source xformSource = new StreamSource(inStream); 
		xformSource.setSystemId(transformFile.toURI().toURL().toExternalForm());
		XsltExecutable executable = compiler.compile(xformSource);
		
		Xslt30Transformer transformer = executable.load30();
		transformer.setErrorListener(errorListener);
		
		MessageListener2 messageListener = new LoggingMessageListener(log);
		transformer.setMessageListener(messageListener);

		Map<QName, XdmValue> parameters = new HashMap<QName, XdmValue>();
		// Assuming that parameters are not namespaced. If they are we'll
		// have to deal with that additional complexity. s
		for (String name : xsltParameters.keySet()) {
			parameters.put(new QName("", name), XdmValue.makeValue(xsltParameters.get(name)));			
		}
		transformer.setStylesheetParameters(parameters);
		
		Source docSource = new StreamSource(docFile);
		
		XdmNode sourceDoc = processor.newDocumentBuilder().build(docSource);
		transformer.setGlobalContextItem(sourceDoc);

		log.info("Applying transform to source document " + docFile.getAbsolutePath() + "...");
	
		@SuppressWarnings("unused")
		XdmValue result = transformer.applyTemplates(docSource);
		log.info("Transform applied.");
	}

	/**
	 * Process a SWPX file to a DOCX file.
	 * @param inFile Single SWPX file
	 * @param outFile If this is a directory, result filename is constructed from input filename. 
	 * @param templateDoc Template DOCX document used when constructing new document
	 * @param log Log to put messages to.
	 */
	public static void handleSingleSwpxDoc(File inFile, File outFile, XWPFDocument templateDoc, Logger log) {
		
		File effectiveOutFile = outFile;
		if (outFile.isDirectory()) {
			String outName = FilenameUtils.getBaseName(inFile.getAbsolutePath()) + ".docx";
			effectiveOutFile = new File(outFile, outName);
		}

    	try {
    		log.info("Generating DOCX file \"" + effectiveOutFile.getAbsolutePath() + "\"");
			if (effectiveOutFile.exists()) {
				if (!effectiveOutFile.delete()) {
					log.error("Could not delete existing DOCX file \"" + effectiveOutFile.getAbsolutePath() + "\". Skipping SWPX file.");
					return;
				}
			}
	    DocxGenerator generator = new DocxGenerator(inFile, effectiveOutFile, templateDoc);
			XmlObject xml = XmlObject.Factory.parse(inFile);

			generator.generate(xml);
			log.info("DOCX file generated.");
		} catch (Throwable e) {
			log.error("Unexpected " + e.getClass().getSimpleName() + ": " + e.getMessage());
			e.printStackTrace();
		}		
	}

	/**
	 * Process all *.swpx files in the input directory, putting the results in the output directory.
	 * <p>NOTE: This method is primarily for testing purposes. During production the SWPX docs are
	 * generated dynamically from the _Book.xml file.</p> 
	 * @param inDir Directory to look for *.swpx files in
	 * @param outDir Directory to write *.docx files to
	 * @param templateDoc Template DOCX document used when constructing new document
	 * @param log Log to write messages to.
	 */
	public static void handleDirectory(File inDir, File outDir, XWPFDocument templateDoc, Logger log) {
		
		FilenameFilter filter = new SuffixFileFilter(".swpx");
		File[] files = inDir.listFiles(filter);
		for (File inFile : files) {
			handleSingleSwpxDoc(inFile, outDir, templateDoc, log);
		}

	}

	/**
	 * Build the command-line options
	 * @return CLI options object ready to use.
	 */
	public static Options buildOptions() {
		Options options = new Options();
    	Option input = Option.builder(OPTION_CHAR_INPUTPATH)
						.required(true)
						.hasArg(true)
						.desc("The path and filename of the Simple WP XML document or directory containing .swpx files.")
						.build();
    	Option output = Option.builder(OPTION_CHAR_OUTPUTPATH)
						.required(true)
						.hasArg(true)
						.desc("The path and filename of the result DOCX file, or directory to contain generated DOCX files")
						.build();
    	Option template = Option.builder(OPTION_CHAR_TEMPLATEPATH)
				.required(true)
				.hasArg(true)
				.desc("The path and filename of the template DOTX file.")
				.build();
    	Option transform = Option.builder(OPTION_CHAR_TRANSFORMPATH)
				.required(false)
				.hasArg(true)
				.desc("The path and filename of the XSLT transform for generating SWPX documents.")
				.build();
      Option dpi = Option.builder("d")
          .longOpt("dpi")
        .required(false)
        .hasArg(true)
        .desc("The dots-per-inch value to use when converting pixels to absolute measurements, e.g., \"72\" or \"96\".")
        .build();
      Option chunkLevel = Option.builder(OPTION_CHAR_CHUNKLEVEL)
          .longOpt("chunklevel")
        .required(false)
        .hasArg(true)
        .desc("Control generation of separate DOCX files from input sections. Default is \"root\", create a single chunk. "
            + "Values are determined by the details of the XSLT transform.")
        .build();
      Option catalog = Option.builder(OPTION_CHAR_CATALOG)
          .longOpt("catalog")
        .required(false)
        .hasArg(true)
        .desc("Semicolon-separated list of XML catalog file names, as for the Saxon -catalog option.")
        .build();
    			
    	options.addOption(input);
    	options.addOption(output);
    	options.addOption(template);
    	options.addOption(transform);
    	options.addOption(dpi);
      options.addOption(chunkLevel);
    	options.addOption(catalog);

		return options;
	}
}
