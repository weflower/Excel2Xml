package will_dot_flowers_at_gmail.excel2xml;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.HashMap;
import java.util.Set;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

////////////////////////////////////////////////////////////////////////////////
//  Project:  Excel2Xis
//  File:     OutputUtils.java
//
//  Name:     Will Flowers
//  Email:    will.flowers@gmail.com
////////////////////////////////////////////////////////////////////////////////

/**
 * Utility methods for outputting an Excel file to one or multiple files.
 *
 * <p/> Bugs: None.
 *
 * @author Will Flowers
 *
 */
public class OutputUtils
{
	
	private static Transformer myTransformer;
	
	/**
	 *
	 * Output a WorkBook as separate XML files based on worksheet. Each file contains
	 * one table element.
	 *
	 * @param myWorkBook
	 * @param document
	 * @throws TransformerException
	 * @throws ParserConfigurationException 
	 */
	protected static void splitMode(WorkBook myWorkBook) throws TransformerException, ParserConfigurationException
	{
		HashMap<String,String> outFiles = Excel2Xml.getOutFiles();
		
		int sheetNum = 0;
		for (WorkSheet thisSheet : myWorkBook)
		{
			// Set up document transformer
			OutputUtils.setMyTransformer();
						
			// Initializing the XML document
			Document document = OutputUtils.buildNewDoc();			
			
			Element rootElement = document.createElement("xml");
			document.appendChild(rootElement);
			
			Element tableElement = document.createElement("table");
			rootElement.appendChild(tableElement);
			
			// Process header
			Element headElement = document.createElement("head");
			tableElement.appendChild(headElement);
			for (String str : thisSheet.getMyHeader())
			{
				Element cellElement = document.createElement("cell");
				headElement.appendChild(cellElement);
				cellElement.appendChild(document.createTextNode(str));
			}

			// Process each row
			for (int i = 1; i < thisSheet.getSize(); i++)
			{
				Element rowElement = document.createElement("row");
				tableElement.appendChild(rowElement);

				// Process each column in a row
				for (String s : thisSheet.getRow(i))
				{
					Element cellElement = document.createElement("cell");
					rowElement.appendChild(cellElement);
					cellElement.appendChild(document.createTextNode(s));
				}
			}

			// Construct sheet file name
			String sheetFileName = null;
			String sheetName = thisSheet.getSheetName();
			if (sheetNum < 9)
			{
				sheetFileName = Excel2Xml.getOutDir() + "/" + "0" + (sheetNum + 1) + "_" + sheetName + ".xml";
			}
			else
			{
				sheetFileName = Excel2Xml.getOutDir() + "/" + (sheetNum + 1) + "_" + sheetName + ".xml";
			}

			// Create XML output
			OutputUtils.writeFile(document, sheetFileName);

			// Add output file to list for post-processing
			outFiles.put(sheetFileName,sheetName);
			sheetNum++;
		}
	}

	/**
	 *
	 * Outputs a WorkBook as one XML file with multiple table elements
	 *
	 * @param xlsPath
	 * @param myWorkBook
	 * @param document
	 * @throws TransformerException
	 * @throws ParserConfigurationException 
	 */
	protected static void fullMode(String xlsPath, WorkBook myWorkBook) throws TransformerException, ParserConfigurationException
	{
		
		HashMap<String,String> outFiles = Excel2Xml.getOutFiles();
		
		// Set up document transformer
		OutputUtils.setMyTransformer();
					
		// Initializing the XML document
		Document document = OutputUtils.buildNewDoc();

		Element rootElement = document.createElement("xml");
		document.appendChild(rootElement);

		// Turn each sheet into its own table in one output file
		for (int sheetNum = 0; sheetNum < myWorkBook.length(); sheetNum++)
		{
			WorkSheet thisSheet = myWorkBook.getSheet(sheetNum);

			Element tableElement = document.createElement("table");
			rootElement.appendChild(tableElement);
			
			// Process header
    		Element headElement = document.createElement("head");
    		tableElement.appendChild(headElement);
    		for (String str : thisSheet.getMyHeader())
    		{
    			Element cellElement = document.createElement("cell");
    			headElement.appendChild(cellElement);
    			cellElement.appendChild(document.createTextNode(str));
    		}

			// Process each row
			for (int i = 1; i < thisSheet.getSize(); i++)
			{
				Element rowElement = document.createElement("row");
				tableElement.appendChild(rowElement);

				// Process each column in a row
				for (String s : thisSheet.getRow(i))
				{
					Element cellElement = document.createElement("cell");
					rowElement.appendChild(cellElement);
					cellElement.appendChild(document.createTextNode(s));
				}
			}
		}

		// Construct output filename from input file
		String sheetName = xlsPath.substring(0,xlsPath.length() - 5); // use file name - ext
		String fileName = Excel2Xml.getOutDir() + "/" + sheetName + ".xml";

		OutputUtils.writeFile(document, fileName);

		// Add output file to map: file name, sheetName
		outFiles.put(fileName,sheetName);
	}
	
	/**
	 * Creates a new Document object
	 *
	 * @return
	 * @throws ParserConfigurationException 
	 */
	private static Document buildNewDoc() throws ParserConfigurationException
	{
		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		DocumentBuilder builder = factory.newDocumentBuilder();
		Document document = builder.newDocument();
		
		return document;
	}
	
	/**
	 * Sets up a Transformer object to change the source to the result doc
	 * 
	 * @param myTransformer
	 * @throws TransformerConfigurationException
	 */
	protected static void setMyTransformer() throws TransformerConfigurationException
	{
		TransformerFactory tFactory = TransformerFactory.newInstance();

		Transformer transformer = tFactory.newTransformer();
		// Add indentation to output
		transformer.setOutputProperty(OutputKeys.INDENT, "yes");
		transformer.setOutputProperty(
				"{http://xml.apache.org/xslt}indent-amount", "2");

		OutputUtils.myTransformer = transformer;
	}

	// Returns the output Transformer
	protected static Transformer getMyTransformer()
	{
		return myTransformer;
	}
	
	/**
	 * 
	 * Writes a Document object to a file based on the outName. 
	 *
	 * @param document
	 * @param outName
	 * @throws TransformerException
	 */
	private static void writeFile(Document document, String outName) throws TransformerException
	{
		// Build output document
		DOMSource source = new DOMSource(document);
		StreamResult result = new StreamResult(new File(outName));
		Transformer transformer = OutputUtils.getMyTransformer();
		transformer.transform(source, result);
	}

	/**
	 *
	 * Post process a list of files with an XSL transform
	 *
	 * @param outFiles
	 * @param stylesheet
	 * @throws FileNotFoundException
	 * @throws TransformerException
	 */
	protected static void postProcess(HashMap<String,String> outFiles, String stylesheet) throws FileNotFoundException, TransformerException
	{
		
		Set<String> myKeys = outFiles.keySet();
		
		for (String fileName : myKeys)
		{
			// Instantiate a TransformerFactory.
			javax.xml.transform.TransformerFactory tFactory =
			                  javax.xml.transform.TransformerFactory.newInstance();

			// Use the TransformerFactory to process the stylesheet Source and generate a Transformer.
			javax.xml.transform.Transformer xslTransformer = tFactory.newTransformer
			                (new javax.xml.transform.stream.StreamSource(stylesheet));
			
			xslTransformer.setParameter("filename", fileName);
			xslTransformer.setParameter("sheetname", outFiles.get(fileName));

			// Use the Transformer to transform an XML Source and send the output to a Result object.
			xslTransformer.transform
			    (new javax.xml.transform.stream.StreamSource(fileName),
			     new javax.xml.transform.stream.StreamResult(new
			                                  java.io.FileOutputStream(fileName + ".out")));
		}
	}
}
