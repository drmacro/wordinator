package com.municode.munipub2docx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

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
import org.apache.xmlbeans.XmlObject;

import com.municode.munipub2docx.generator.DocxGeneratingOutputUriResolver;
import com.municode.munipub2docx.generator.DocxGenerator;

import net.sf.saxon.lib.FeatureKeys;
import net.sf.saxon.lib.OutputURIResolver;
import net.sf.saxon.s9api.Processor;
import net.sf.saxon.s9api.QName;
import net.sf.saxon.s9api.XdmValue;
import net.sf.saxon.s9api.Xslt30Transformer;
import net.sf.saxon.s9api.XsltCompiler;
import net.sf.saxon.s9api.XsltExecutable;

/**
 * Command-line application to generate DOCX files from
 * various inputs.
 *
 */
public class MakeDocx 
{
	public static void main( String[] args ) throws ParseException
    {
    	Options options = buildOptions();
    	CommandLineParser parser = new DefaultParser();
    	CommandLine cmd = parser.parse( options, args);
    	
    	String inDocPath = cmd.getOptionValue("i");
    	String docxPath = cmd.getOptionValue("o");
    	String templatePath = cmd.getOptionValue("t");
    	String transformPath = cmd.getOptionValue("x");
    	String chunkLevel = cmd.getOptionValue("c");
    	chunkLevel = chunkLevel == null ? "section" : chunkLevel;
    	
    	// FIXME: Set up proper Java logging.
    	System.out.println("+ [INFO] Input document or directory='" + inDocPath + "'");
    	System.out.println("+ [INFO] Output directory           ='" + docxPath + "'");
    	System.out.println("+ [INFO] DOTX template              ='" + templatePath + "'");
    	System.out.println("+ [INFO] XSLT template              =" + (transformPath == null ? "Not specified" : "'" + transformPath + "'"));
    	System.out.println("+ [INFO] Chunk level                ='" + chunkLevel + "'");
    	
    	// Check that the input file exists.
    	// For now, always overwriting the DOCX file without confirmation.
    	
    	File inFile = new File(inDocPath);
    	if (!inFile.exists()) {
    		System.err.println("- [ERROR] Input file '" + inFile.getAbsolutePath() + "' not found. Cannot continue."); 
    		System.exit(1);
    	}
    	File templateFile = new File(templatePath);
    	if (!templateFile.exists()) {
    		System.err.println("- [ERROR] Template file '" + templateFile.getAbsolutePath() + "' not found. Cannot continue."); 
    		System.exit(1);
    	}
    	File outFile = new File(docxPath);
    	
    	File outDir = outFile; // Normal case: specify output directory
    	if (outFile.getName().endsWith(".docx")) {
    		outDir = outFile.getParentFile();
    	}
    	
    	if (!outDir.exists()) {
    		System.out.println("Making output directory '" + outDir.getAbsolutePath() + "'...");
    		if (!outDir.mkdirs()) {
    			System.err.println("- [ERROR] Failed to create output directory '" + outDir.getAbsolutePath() + "'. Cannot continue");
        		System.exit(1);
    		}
    	}
    	
    	File transformFile = null;
    	if (null != transformPath) {
    		transformFile = new File(transformPath);
        	if (!transformFile.exists()) {
        		System.err.println("- [ERROR] XSLT transform file '" + transformFile.getAbsolutePath() + "' not found. Cannot continue."); 
        		System.exit(1);
        	}
    	}
    	
    	try {
	    	if (inFile.isDirectory()) {
	    		File cand = new File(inFile, "_Book.xml");
	    		if (cand.exists()) {
	    			processHtmlBook(cand, outDir, templateFile, transformFile, chunkLevel);
	    		} else {
	        		handleDirectory(inFile, outDir, templateFile);    			
	    		}
	    	} else {
	    		if (inFile.getName().equalsIgnoreCase("_book.xml")) {
	    			processHtmlBook(inFile, outDir, templateFile, transformFile, chunkLevel);
	    		} else {
	    			handleSingleFile(inFile, outFile, templateFile);
	    		}
	    	}
    	} catch (Exception e) {
    		System.out.println("- [ERROR] " + e.getClass().getSimpleName() + ": " + e.getMessage());
    		System.exit(1);
    	}
    	
    	
    }

	/**
	 * Process a _Book.xml file to a set of DOCX files
	 * @param bookFile the _Book.xml file to process
	 * @param outDir Directory to put the DOCX files in
	 * @param templateFile DOTX file to use in constructing new DOCX files.
	 * @param transformFile The file containing the XSLT transform for generating SWPX documents
	 * @param chunkLevel 
	 * @throws Exception 
	 */
	private static void processHtmlBook(
			File bookFile, 
			File outDir, 
			File templateFile, 
			File transformFile, 
			String chunkLevel) throws Exception {
		// Apply transform to book file to generate Simple WP XML documents
		
		if (transformFile == null) {
			throw new RuntimeException("-x (transform) parameter not specified. If the input is a _Book.xml file, you must specify the -x parameter");
		}
		
		Processor processor = new Processor(false);
		OutputURIResolver outputResolver = new DocxGeneratingOutputUriResolver(outDir, templateFile);
		processor.setConfigurationProperty(FeatureKeys.OUTPUT_URI_RESOLVER, outputResolver);
		
		// FIXME: Set up proper logger. See 
		// https://www.saxonica.com/html/documentation/using-xsl/embedding/s9api-transformation.html
		XsltCompiler compiler = processor.newXsltCompiler();
		
		InputStream inStream = new FileInputStream(transformFile);
		Source xformSource = new StreamSource(inStream); 
		xformSource.setSystemId(transformFile.toURI().toURL().toExternalForm());
		XsltExecutable executable = compiler.compile(xformSource);
		
		Xslt30Transformer transformer = executable.load30();
		Map<QName, XdmValue> parameters = new HashMap<QName, XdmValue>();
		parameters.put(new QName("", "chunk-level"), XdmValue.makeValue(chunkLevel));
		transformer.setStylesheetParameters(parameters);
		
		Source bookSource = new StreamSource(bookFile);
		System.out.println("+ [INFO] Applying transform to source document " + bookFile.getAbsolutePath() + "...");
	
		@SuppressWarnings("unused")
		XdmValue result = transformer.applyTemplates(bookSource);
		System.out.println("Transform applied.");
		// Direct result is the generation log XML document
		// At this point, all the DOCX files should be generated.
		// System.out.println(result.iterator().next().getStringValue());
	}

	/**
	 * Process a SWPX file to a DOCX file.
	 * @param inFile Single SWPX file
	 * @param outFile If this is a directory, result filename is constructed from input filename. 
	 * @param templateFile DOTX file to use as a template when constructing the new document.
	 */
	private static void handleSingleFile(File inFile, File outFile, File templateFile) {
		
		File effectiveOutFile = outFile;
		if (outFile.isDirectory()) {
			String outName = FilenameUtils.getBaseName(inFile.getAbsolutePath()) + ".docx";
			effectiveOutFile = new File(outFile, outName);
		}

    	try {
    		System.out.println("+ [INFO] Generating DOCX file \"" + effectiveOutFile.getAbsolutePath() + "\"");
			if (effectiveOutFile.exists()) {
				if (!effectiveOutFile.delete()) {
					System.err.println("- [ERROR] Could not delete existing DOCX file \"" + effectiveOutFile.getAbsolutePath() + "\". Skipping SWPX file.");
					return;
				}
			}
	    	DocxGenerator generator = new DocxGenerator(inFile, effectiveOutFile, templateFile);
			XmlObject xml = XmlObject.Factory.parse(inFile);

			generator.generate(xml);
    		System.out.println("+ [INFO] DOCX file generated.");
		} catch (Throwable e) {
			System.err.println("- [ERROR] Unexpected " + e.getClass().getSimpleName() + ": " + e.getMessage());
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
	 */
	private static void handleDirectory(File inDir, File outDir, File templateFile) {
		
		FilenameFilter filter = new SuffixFileFilter(".swpx");
		File[] files = inDir.listFiles(filter);
		for (File inFile : files) {
			handleSingleFile(inFile, outDir, templateFile);
		}

	}

	/**
	 * Build the command-line options
	 * @return CLI options object ready to use.
	 */
	private static Options buildOptions() {
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
    	Option chunkLevel = Option.builder("c")
				.required(false)
				.hasArg(true)
				.desc("The chunking level, one of \"chapter\" or \"section\"")
				.build();
    			
    	options.addOption(input);
    	options.addOption(output);
    	options.addOption(template);
    	options.addOption(transform);
    	options.addOption(chunkLevel);
		return options;
	}
}
