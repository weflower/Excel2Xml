Author: Will Flowers (will.flowers@gmail.com)

Usage: java -jar excel2xml.jar -f fileName [-split] [-xsl stylesheet]

The options are:
-f file name
[-split] - Split the output into separate files by worksheet
[-xsl stylesheet] - An optional XSL stylesheet to post-process the output files.

This Java application transforms an Excel spreadsheet (.xlsx, Excel 2007+) into an XML file.
Uses the Apache POI libraries for Java (https://poi.apache.org/).

The base XML output has the format of:
   xml
      table
      	head
      		cell
        row
			cell
			
You can optionally add a stylesheet to post-process these files. Files are 
output to the ./output directory. Post-processed files have an additional .out
extension.


