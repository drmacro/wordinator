package org.wordinator.xml2docx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import javax.xml.transform.ErrorListener;
import javax.xml.transform.Source;
import javax.xml.transform.stream.StreamSource;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.filefilter.SuffixFileFilter;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.xmlbeans.XmlObject;
import org.wordinator.xml2docx.generator.DocxGeneratingOutputUriResolver;
import org.wordinator.xml2docx.generator.DocxGenerator;
import org.wordinator.xml2docx.saxon.Log4jSaxonLogger;
import org.wordinator.xml2docx.saxon.LoggingMessageListener;

import net.sf.saxon.lib.FeatureKeys;
import net.sf.saxon.lib.OutputURIResolver;
import net.sf.saxon.lib.StandardErrorListener;
import net.sf.saxon.lib.StandardLogger;
import net.sf.saxon.s9api.MessageListener;
import net.sf.saxon.s9api.Processor;
import net.sf.saxon.s9api.QName;
import net.sf.saxon.s9api.XdmValue;
import net.sf.saxon.s9api.Xslt30Transformer;
import net.sf.saxon.s9api.XsltCompiler;
import net.sf.saxon.s9api.XsltExecutable;

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
	
	public static final Logger log = LogManager.getLogger(MakeDocx.class.getSimpleName());
			
	public static final String XSLT_PARAM_CHUNKLEVEL = "chunklevel";

	public static void main( String[] args ) throws ParseException
    {
    	Options options = buildOptions();
    	CommandLineParser parser = new DefaultParser();
    	CommandLine cmd = parser.parse( options, args);
    	
    	Map<String, String> xsltParameters = new HashMap<String, String>();

    	
    	handleCommandLine(cmd, xsltParameters, log);    	
    	
    }

	/**
	 * Does the actual command line processing. You can call this from your own
	 * command line processor if you need additional command-line options, for example,
	 * to set additional XSLT parameters.
	 * @param cmd The command line command.
	 * @param xsltParameters XSLT parameters to use. Note that the chunklevel parameter will be set automatically.
	 * @param log Logger to log messages to.
	 */
	public static void handleCommandLine(
			CommandLine cmd, 
			Map<String, String> xsltParameters, 
			Logger log) {
		String inDocPath = cmd.getOptionValue("i");
    	String docxPath = cmd.getOptionValue("o");
    	String templatePath = cmd.getOptionValue("t");
    	String transformPath = cmd.getOptionValue("x");
    	String chunkLevel = cmd.getOptionValue("c");
    	chunkLevel = chunkLevel == null ? "root" : chunkLevel;
    	
    	// FIXME: Set up proper Java logging.
    	log.info("Input document or directory='" + inDocPath + "'");
    	log.info("Output directory           ='" + docxPath + "'");
    	log.info("DOTX template              ='" + templatePath + "'");
    	log.info("XSLT template              =" + (transformPath == null ? "Not specified" : "'" + transformPath + "'"));
    	log.info("Chunk level                ='" + chunkLevel + "'");
    	
    	// Check that the input file exists.
    	// For now, always overwriting the DOCX file without confirmation.
    	
    	File inFile = new File(inDocPath);
    	if (!inFile.exists()) {
    		log.error("Input file '" + inFile.getAbsolutePath() + "' not found. Cannot continue."); 
    		System.exit(1);
    	}
    	File templateFile = new File(templatePath);
    	if (!templateFile.exists()) {
    		log.error("Template file '" + templateFile.getAbsolutePath() + "' not found. Cannot continue."); 
    		System.exit(1);
    	}
    	File outFile = new File(docxPath);
    	
    	File outDir = outFile; // Normal case: specify output directory
    	if (outFile.getName().endsWith(".docx")) {
    		outDir = outFile.getParentFile();
    	}
    	
    	if (!outDir.exists()) {
    		log.info("Making output directory '" + outDir.getAbsolutePath() + "'...");
    		if (!outDir.mkdirs()) {
    			log.error("Failed to create output directory '" + outDir.getAbsolutePath() + "'. Cannot continue");
        		System.exit(1);
    		}
    	}
    	
    	File transformFile = null;
    	if (null != transformPath) {
    		transformFile = new File(transformPath);
        	if (!transformFile.exists()) {
        		log.error("XSLT transform file '" + transformFile.getAbsolutePath() + "' not found. Cannot continue."); 
        		System.exit(1);
        	}
        	if (!xsltParameters.containsKey(XSLT_PARAM_CHUNKLEVEL)) {
        		xsltParameters.put(XSLT_PARAM_CHUNKLEVEL, chunkLevel);
        	}
    	}
    	
    	try {
    		if (inFile.isDirectory()) {
    			// Assume directory contains *.swpx files 
    			handleDirectory(inFile, outDir, templateFile, log);
    		} else { 
    			if (inFile.getName().endsWith(".swpx")) {
	    			handleSingleSwpxDoc(inFile, outFile, templateFile, log);
	    		} else {
	    			transformXml(inFile, outDir, templateFile, transformFile, xsltParameters, log);
	    		}
    		}
    	} catch (Exception e) {
    		log.error(e.getClass().getSimpleName() + ": " + e.getMessage(), e);
    		System.exit(1);
    	}
	}

	/**
	 * Process an XML document to a set of DOCX files
	 * @param docFile the root XML document to process
	 * @param outDir Directory to put the DOCX files in
	 * @param templateFile DOTX file to use in constructing new DOCX files.
	 * @param transformFile The file containing the XSLT transform for generating SWPX documents
	 * @param xsltParameters Map of parameter names to values to be passed to the XSLT transform.
	 * @param log Log to write messages to.
	 * @throws Exception Any kind of error
	 */
	public static void transformXml(
			File docFile, 
			File outDir, 
			File templateFile, 
			File transformFile, 
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
		OutputURIResolver outputResolver = new DocxGeneratingOutputUriResolver(outDir, templateFile, log);
		processor.setConfigurationProperty(FeatureKeys.OUTPUT_URI_RESOLVER, outputResolver);
		
		// FIXME: Set up proper logger. See 
		// https://www.saxonica.com/html/documentation/using-xsl/embedding/s9api-transformation.html
		XsltCompiler compiler = processor.newXsltCompiler();
		
		InputStream inStream = new FileInputStream(transformFile);
		Source xformSource = new StreamSource(inStream); 
		xformSource.setSystemId(transformFile.toURI().toURL().toExternalForm());
		XsltExecutable executable = compiler.compile(xformSource);
		
		Xslt30Transformer transformer = executable.load30();
		transformer.setErrorListener(errorListener);
		
		MessageListener messageListener = new LoggingMessageListener(log);
		transformer.setMessageListener(messageListener);

		Map<QName, XdmValue> parameters = new HashMap<QName, XdmValue>();
		// Assuming that parameters are not namespaced. If they are we'll
		// have to deal with that additional complexity. s
		for (String name : xsltParameters.keySet()) {
			parameters.put(new QName("", name), XdmValue.makeValue(xsltParameters.get(name)));			
		}
		transformer.setStylesheetParameters(parameters);
		
		Source docSource = new StreamSource(docFile);
		log.info("Applying transform to source document " + docFile.getAbsolutePath() + "...");
	
		@SuppressWarnings("unused")
		XdmValue result = transformer.applyTemplates(docSource);
		log.info("Transform applied.");
	}

	/**
	 * Process a SWPX file to a DOCX file.
	 * @param inFile Single SWPX file
	 * @param outFile If this is a directory, result filename is constructed from input filename. 
	 * @param templateFile DOTX file to use as a template when constructing the new document.
	 * @param log 
	 */
	public static void handleSingleSwpxDoc(File inFile, File outFile, File templateFile, Logger log) {
		
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
	    	DocxGenerator generator = new DocxGenerator(inFile, effectiveOutFile, templateFile);
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
	 * @param templateFile DOTX file to use as template for new DOCX files.
	 * @param log Log to write messages to.
	 */
	public static void handleDirectory(File inDir, File outDir, File templateFile, Logger log) {
		
		FilenameFilter filter = new SuffixFileFilter(".swpx");
		File[] files = inDir.listFiles(filter);
		for (File inFile : files) {
			handleSingleSwpxDoc(inFile, outDir, templateFile, log);
		}

	}

	/**
	 * Build the command-line options
	 * @return CLI options object ready to use.
	 */
	public static Options buildOptions() {
		Options options = new Options();
    	Option input = Option.builder("i")
						.required(true)
						.hasArg(true)
						.desc("The path and filename of the Simple WP XML document or directory containing .swpx files.")
						.build();
    	Option output = Option.builder("o")
						.required(true)
						.hasArg(true)
						.desc("The path and filename of the result DOCX file, or directory to contain generated DOCX files")
						.build();
    	Option template = Option.builder("t")
				.required(true)
				.hasArg(true)
				.desc("The path and filename of the template DOTX file.")
				.build();
    	Option transform = Option.builder("x")
				.required(false)
				.hasArg(true)
				.desc("The path and filename of the XSLT transform for generating SWPX documents.")
				.build();
    			
    	options.addOption(input);
    	options.addOption(output);
    	options.addOption(template);
    	options.addOption(transform);
		return options;
	}
}
