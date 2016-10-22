package will_dot_flowers_at_gmail.excel2xml;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

////////////////////////////////////////////////////////////////////////////////
//  Project:  Excel2Xml
//  File:     Excel2Xml.java
//
//  Name:     Will Flowers
//  Email:    will.flowers@gmail.com
////////////////////////////////////////////////////////////////////////////////

/**
 * Transforms an Excel spreadsheet (.xlsx, Excel 2007+) into an XML file.
 * Uses the Apache POI libraries for Java (https://poi.apache.org/).
 *
 * The base XML output has the format of:
 *    xml
 *       table
 *       	head
 *       		cell
 *          row
 *				cell
 *
 *
 * You can optionally add a stylesheet to post-process these files (only supports native Xalan).
 *
 * Files are output to the ./output directory. Post-processed files have an extra .out
 * extension.
 *
 * Options are:
 * -f file name
 * [-split] - Split the output into separate files by worksheet
 * [-xsl stylesheet] - An optional XSL stylesheet to post-process the output files.
 *
 * <p/>
 * Bugs: None.
 *
 * @author Will Flowers
 *
 */
public class Excel2Xml
{
	private static HashMap<String,String> outFiles; // file name, sheet name
	private static final String OUTDIR = "./output";

	public static void main(String[] args) throws FileNotFoundException, TransformerException
	{
		outFiles = new HashMap<String,String>();

		// If no params, output usage and quit
		if (args.length == 0)
		{
			// Error
			System.out.println("Usage: java -jar excel2xml.jar -f fileName [-split]\n"
					+ "\t[-xsl stylesheet]");
			System.exit(0);
		}
		else
		{
			// Init parameter variables
			String xlsPath = null, mode = null, stylesheet = null;

			// Loop through parameters and set
			int i = 0;
			for (i = 0; i < args.length; i++)
			{
				if ("-f".equalsIgnoreCase(args[i]))
				{
					xlsPath = args[i+1];
				}

				if ("-split".equalsIgnoreCase(args[i]))
				{
					mode = "split";
				}

				if ("-xsl".equalsIgnoreCase(args[i]))
				{
					stylesheet = args[i+1];
				}
			}

			// Set file name and check for extension
			xlsPath = args[1];
			if (!xlsPath.contains(".xlsx"))
			{
				System.err.println("Error: File name must have .xlsx extension.");
				System.exit(0);
			}

			Excel2Xml thisRun = new Excel2Xml();
			thisRun.processExcel(xlsPath, mode);

			// Post process XSL
			if (stylesheet != null)
			{
				OutputUtils.postProcess(outFiles, stylesheet);
			}
		}
	}

	/**
	 *
	 * Constructs a new Excel2Xml object.
	 *
	 */
	public Excel2Xml()	{ /* do nothing */	}

	/**
	 *
	 * Processes an Excel file into XML
	 *
	 * @param xlsPath
	 * @param mode
	 */
	public void processExcel(String xlsPath, String mode)
	{
		InputStream inputStream = null;
		try
		{
			inputStream = new FileInputStream(xlsPath);
		}
		catch (FileNotFoundException e)
		{
			System.err.println("Error: File not found in the specified path.");
			e.printStackTrace();
		}

		try
		{
			// Create  XSSFWorkBook from input file
			XSSFWorkbook xlsWorkBook = new XSSFWorkbook(inputStream);

			// Create new WorkBook object from XSSFWorkBook object
			WorkBook myWorkBook = new WorkBook(xlsWorkBook);

			// Create XML output folder
			new File(Excel2Xml.getOutDir()).mkdir();

			// Output based on mode
			if ("split".equalsIgnoreCase(mode)) // Split each WorkSheet into it's own file
			{
				OutputUtils.splitMode(myWorkBook);
			}
			else //output as one file
			{
				OutputUtils.fullMode(xlsPath, myWorkBook);
			}

			xlsWorkBook.close();
		}
		catch (IOException e)
		{
			System.out.println("IOException " + e.getMessage());
		}
		catch (ParserConfigurationException e)
		{
			System.out.println("ParserConfigurationException " + e.getMessage());
		}
		catch (TransformerConfigurationException e)
		{
			System.out.println("TransformerConfigurationException " + e.getMessage());
		}
		catch (TransformerException e)
		{
			System.out.println("TransformerException " + e.getMessage());
		}
	}

	// Returns the outFiles ArrayList
	public static HashMap<String,String> getOutFiles()
	{
		return outFiles;
	}

	// Returns output directory
	public static String getOutDir()
	{
		return OUTDIR;
	}
}
