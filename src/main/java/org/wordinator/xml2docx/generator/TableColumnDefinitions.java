package org.wordinator.xml2docx.generator;

import java.util.ArrayList;
import java.util.List;

/*
 * Manages the column definitions for a table.
 */
public class TableColumnDefinitions {
	
	private List<TableColumnDefinition> colDefs = new ArrayList<TableColumnDefinition>();

	/**
	 * Construct a new column definition, making it the last column
	 * in the list of columns.
	 * @return New column definition
	 */
	public TableColumnDefinition newColumnDef() {
		TableColumnDefinition colDef = new TableColumnDefinition(this);
		this.colDefs.add(colDef);
		return colDef;
	}

	/**
	 * Get the list of column definitions
	 * @return List of column definitions
	 */
	public List<TableColumnDefinition> getColumnDefinitions() {
		return this.colDefs;
	}

	/**
	 * Get the column definition for the specified column.
	 * @param columnIndex Zero-index column number
	 * @return The column definition for the specified column. If there is no 
	 * existing definition a new one is created.
	 */
	public TableColumnDefinition get(int columnIndex) {
		TableColumnDefinition def = columnIndex >= colDefs.size() ? newColumnDef() : colDefs.get(columnIndex); 
		return def;
	}

}
