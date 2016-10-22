package will_dot_flowers_at_gmail.excel2xml;
import java.util.ArrayList;
import java.util.Iterator;

////////////////////////////////////////////////////////////////////////////////
//  Course:   CSC 249 Fall 2016
//  Section:  0001
//
//  Project:  Excel2Xml
//  File:     TableRows.java
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
public class TableRow implements Iterable<String>
{
	ArrayList<String> myCols;

	public TableRow()
	{
		myCols = new ArrayList<String>();
	}

	public void addCol(String s)
	{
		myCols.add(s);
	}

	public String getCol(int idx)
	{
		return myCols.get(idx);
	}
	
	public String toString()
	{
		String str = "";
		
		for (String s : this.myCols)
		{
			str += s + ",";
		}
		
		return str;
	}

	/* (non-Javadoc)
	 * @see java.lang.Iterable#iterator()
	 */
	@Override
	public Iterator<String> iterator() {
        Iterator<String> it = new Iterator<String>() {

            private int currentIndex = 0;

            @Override
            public boolean hasNext() {
                return currentIndex < myCols.size() && myCols.get(currentIndex) != null;
            }

            @Override
            public String next() {
                return myCols.get(currentIndex++);
            }

            @Override
            public void remove() {
                throw new UnsupportedOperationException();
            }
        };
        return it;
    }
}
