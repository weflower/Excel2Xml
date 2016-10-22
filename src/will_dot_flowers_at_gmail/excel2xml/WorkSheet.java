package will_dot_flowers_at_gmail.excel2xml;

import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

////////////////////////////////////////////////////////////////////////////////
//  Course:   CSC 249 Fall 2016
//  Section:  0001
//
//  Project:  Excel2Xml
//  File:     WorkSheet.java
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
public class WorkSheet implements Iterable<TableRow>
{
	ArrayList<TableRow> myRows;
	String sheetName;
	TableRow myHeader;

	/**
	 *
	 * Constructs a new WorkSheet object from an XSSFsheet
	 *
	 * @param sheet
	 */
	public WorkSheet(XSSFSheet sheet)
	{

		sheetName = sheet.getSheetName();
		myRows = new ArrayList<TableRow>();

		// Process the header row
		XSSFRow headRow = (XSSFRow) sheet.getRow(sheet.getFirstRowNum());
		myHeader = new TableRow();
		for (Cell c : headRow)
		{
			myHeader.addCol(doCells(c));
		}
		
		// Process rest of table
		for (Row rawRow : sheet)
		{
			XSSFRow row = (XSSFRow) rawRow;

			// get a row, iterate through cells.
			Iterator<Cell> cells = row.cellIterator();
			TableRow rowData = new TableRow();

			while (cells.hasNext())
			{
				rowData.addCol(doCells(cells.next()));

			} // end while

			// Add full row
			myRows.add(rowData);
		}
	}

	/**
	 *
	 * Add a new row to this WorkSheet
	 *
	 * @param r
	 */
	public void addRow(TableRow r)
	{
		myRows.add(r);
	}

	@SuppressWarnings("deprecation")
	protected String doCells(Cell c)
	{

		if (c != null)
		{
			XSSFCell cell = (XSSFCell) c;
			// System.out.println ("Cell : " + cell.getCellNum ());
			switch (cell.getCellType())
			{
				case XSSFCell.CELL_TYPE_NUMERIC:
				{
					// NUMERIC CELL TYPE
					// outWrite.println("Numeric: " + cell.getNumericCellValue());
					return cell.getNumericCellValue() + "";
				}
				case XSSFCell.CELL_TYPE_STRING:
				{
					// STRING CELL TYPE
					XSSFRichTextString richTextString = cell.getRichStringCellValue();

					// outWrite.println("String: " + richTextString.getString());
					return richTextString.getString();
				}
				default:
				{
					// types other than String and Numeric.
					// System.out.println("Cell type not supported. " + "Sheet: " + cell.getSheet().getSheetName() + ", Row: " + rowNumber + ", Cell: " + cell.getColumnIndex());
					return "";
				}
			}

		}
		else
			return null;
	}

	/**
	 *
	 * Return a row for this WorkSheet
	 *
	 * @param idx
	 * @return
	 */
	public TableRow getRow(int idx)
	{
		return myRows.get(idx);
	}

	public String getSheetName()
	{
		return sheetName;
	}

	public int getSize()
	{
		return myRows.size();
	}
	
	public TableRow getMyHeader()
	{
		return myHeader;
	}
	
	public String toString()
	{
		String str = "";
		
		str += "Sheet name: " + this.getSheetName() + "\n";
		
		for (TableRow r : this)
		{
			str += r.toString() + "\n";
		}
		
		return str;
	}

	/* (non-Javadoc)
	 * @see java.lang.Iterable#iterator()
	 */
	@Override
	public Iterator<TableRow> iterator() {
        Iterator<TableRow> it = new Iterator<TableRow>() {

            private int currentIndex = 0;

            @Override
            public boolean hasNext() {
                return currentIndex < myRows.size() && myRows.get(currentIndex) != null;
            }

            @Override
            public TableRow next() {
                return myRows.get(currentIndex++);
            }

            @Override
            public void remove() {
                throw new UnsupportedOperationException();
            }
        };
        return it;
    }
}
