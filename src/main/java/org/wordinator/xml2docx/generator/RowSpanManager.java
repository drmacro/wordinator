package org.wordinator.xml2docx.generator;

import java.util.HashMap;
import java.util.Map;

/**
 * Manage setting of vertical spanning across rows in a table.
 * <p>Tracks the number of rows that have participated in a given
 * column's current span.</p>
 *
 */
public class RowSpanManager {

	private Map<Integer, Integer> columns = new HashMap<Integer, Integer>();

	public void addColumn(int cellCtr, int spanval) {
		this.columns.put(new Integer(cellCtr), new Integer(spanval));
		// Account for the first cell of the span.
		includeCell(cellCtr);
	}

	/**
	 * Count the current cell as being included the column's vertical
	 * span
	 * @param cellCtr Column number of the cell participating in the vertical span
	 * @return Remaining rows to be spanned or -1 if the cell is not found (or too many have been counted)
	 */
	public int includeCell(int cellCtr) {
		Integer counter = columns.get(Integer.valueOf(cellCtr));
		if (counter != null) {
			counter--;
			columns.put(Integer.valueOf(cellCtr), counter);
			return counter.intValue();
		}
		return -1;
	}

}
