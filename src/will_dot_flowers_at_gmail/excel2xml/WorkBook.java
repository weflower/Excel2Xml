package will_dot_flowers_at_gmail.excel2xml;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

////////////////////////////////////////////////////////////////////////////////
//  Course:   CSC 249 Fall 2016
//  Section:  0001
//
//  Project:  Excel2Xml
//  File:     WorkBook.java
//
//  Name:     Will Flowers
//  Email:    wflowers@my.waketech.edu
////////////////////////////////////////////////////////////////////////////////

/**
 * (Insert a comment that briefly describes the purpose of this class definition.)
 *
 * <p/> Bugs: None.
 *
 * @author Will Flowers
 *
 */
public class WorkBook implements Iterable<WorkSheet>
{
	WorkSheet[] mySheets;

	/**
	 *
	 * Constructs a new WorkBook object from an XSSFWorkBook.
	 *
	 * @param xlsWorkBook
	 */
	public WorkBook(XSSFWorkbook xlsWorkBook)
	{
		mySheets = new WorkSheet[xlsWorkBook.getNumberOfSheets()];

		for (int currSheet = 0; currSheet < xlsWorkBook.getNumberOfSheets(); currSheet++)
		{
			// Create a new workSheet based on the XSSFWorkSheet
			WorkSheet thisSheet = new WorkSheet(xlsWorkBook.getSheetAt(currSheet));

			// Add the WorkSheet we created to our book
			mySheets[currSheet] = thisSheet;
		}
	}

	public int length()
	{
		return this.mySheets.length;
	}

	/**
	 *
	 * Gets a WorkSheet in this book based on the index.
	 *
	 * @param idx
	 * @return
	 */
	public WorkSheet getSheet(int idx)
	{
		return mySheets[idx];
	}

	/* (non-Javadoc)
	 * @see java.lang.Iterable#iterator()
	 */
	@Override
	public Iterator<WorkSheet> iterator() {
        Iterator<WorkSheet> it = new Iterator<WorkSheet>() {

            private int currentIndex = 0;

            @Override
            public boolean hasNext() {
                return currentIndex < mySheets.length && mySheets[currentIndex] != null;
            }

            @Override
            public WorkSheet next() {
                return mySheets[currentIndex++];
            }

            @Override
            public void remove() {
                throw new UnsupportedOperationException();
            }
        };
        return it;
    }
}
