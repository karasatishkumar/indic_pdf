package ooo.connector.example;


import java.util.ArrayList;

import ooo.connector.BootstrapSocketConnector;


import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.connection.NoConnectException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.table.BorderLine;
import com.sun.star.table.TableBorder;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.text.TableColumnSeparator;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextTable;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

public class IndicPDFGenerator {

	private static final String OOO_EXEC_FOLDER = "C:/Program Files/OpenOffice.org 2.4/program/";

	private static final String TEMPLATE_FOLDER = "C:/OpenOffice_1_4/";

	private static final String FILE_URL_PREFIX = "file:///";

	private static final String PDF_DOCUMENT_EXTENSION = ".pdf";
	
	/**
	 * Converts an OOo text document (.odt) to a PDF file using a
	 * BootstrapConnector.
	 * 
	 * @param args
	 *            The command line arguments
	 */
	public static void main(String[] args) {

		try {

			IndicPDFGenerator indicPDFGenerator = new IndicPDFGenerator();
			XComponentContext xContext = indicPDFGenerator.getContext();
			XComponent xWriterComponent = indicPDFGenerator
					.getWriterComponent(xContext);

			XTextDocument xTextDocument = indicPDFGenerator
					.openWriter(xWriterComponent);

			XTextTable xTextTable = indicPDFGenerator.CreateTable(
					xTextDocument, xWriterComponent);

			indicPDFGenerator.setTableProperties(xTextTable);

			XCellRange xCellRange = (XCellRange) UnoRuntime.queryInterface(
					XCellRange.class, xTextTable);

			
			
			indicPDFGenerator.fillFirstRow(xCellRange);
			indicPDFGenerator.CreateTableRow(xCellRange, indicPDFGenerator.getText("hi"));

			

			// CreateTableRow(xCellRange);
			indicPDFGenerator.storePDFComponent(xWriterComponent,
					FILE_URL_PREFIX + TEMPLATE_FOLDER
							+ "test" + "_TEMP"
							+ PDF_DOCUMENT_EXTENSION);
			
		} catch (NoConnectException e) {
			System.out.println("OOo is not responding");
			e.printStackTrace();
		} catch (java.lang.Exception e) {
			e.printStackTrace();
		} finally {
			System.exit(0);
		}
	}
	
	

	/**
	 * This method returns the object of XComponentLoader from the XComponent
	 * Context
	 * 
	 * @param remoteContext -
	 *            XComponentContext object
	 * @return - XComponentLoader object
	 * @throws Exception -
	 *             Checked exception
	 */
	private XComponentLoader getComponentLoader(XComponentContext remoteContext)
			throws Exception {

		XMultiComponentFactory remoteServiceManager = remoteContext
				.getServiceManager();
		Object desktop = remoteServiceManager.createInstanceWithContext(
				"com.sun.star.frame.Desktop", remoteContext);
		XComponentLoader xcomponentloader = (XComponentLoader) UnoRuntime
				.queryInterface(XComponentLoader.class, desktop);

		return xcomponentloader;
	}

	private XTextDocument openWriter(XComponent xWriterComponent) {
		// define variables
		XTextDocument xDoc = null;

		try {
			xDoc = (XTextDocument) UnoRuntime.queryInterface(
					XTextDocument.class, xWriterComponent);

		} catch (java.lang.Exception e) {
			System.err.println(" Exception " + e);
			e.printStackTrace(System.err);
		}
		return xDoc;
	}

	/**
	 * This method is responsible fot getting a XComponent.
	 * 
	 * @param xContext -
	 *            XComponentContext object
	 * @return - XComponent object
	 */
	private XComponent getWriterComponent(XComponentContext xContext) {
		// define variables
		XComponentLoader xCLoader;
		XComponent xComp = null;

		try {
			xCLoader = getComponentLoader(xContext);
			PropertyValue[] szEmptyArgs = setDocumentProperties();

			String strDoc = "private:factory/swriter";
			xComp = xCLoader.loadComponentFromURL(strDoc, "_blank", 0,
					szEmptyArgs);

		} catch (Exception e) {
			System.err.println(" Exception " + e);
			e.printStackTrace(System.err);
		}
		return xComp;
	}

	/**
	 * Need to remove from the class. No use of this method in e-Stamp
	 * application. Created to demonstrate the text rendering.
	 * 
	 * @param xText
	 * @throws com.sun.star.uno.Exception
	 */
	protected static void manipulateText(XText xText)
			throws com.sun.star.uno.Exception {
		// simply set whole text as one string
		/*
		 * xText.setString("He lay flat on the brown, pine-needled floor of the
		 * forest, " + "his chin on his folded arms, and high overhead the wind
		 * blew in the tops " + "of the pine trees."); // create text cursor for
		 * selecting and formatting XTextCursor xTextCursor =
		 * xText.createTextCursor(); XPropertySet xCursorProps =
		 * (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class,
		 * xTextCursor); // use cursor to select "He lay" and apply bold italic
		 * xTextCursor.gotoStart(false); xTextCursor.goRight((short)6, true); //
		 * from CharacterProperties xCursorProps.setPropertyValue("CharPosture",
		 * com.sun.star.awt.FontSlant.ITALIC);
		 * xCursorProps.setPropertyValue("CharWeight", new
		 * Float(com.sun.star.awt.FontWeight.BOLD)); // add more text at the end
		 * of the text using insertString xTextCursor.gotoEnd(false);
		 */
		xText
				.setString("Copyright (C) 1991, 1999 Free Software Foundation, Inc..\n "
						+ "\n"
						+ "This software is made to generate indic language .pdf files.\n"
						+ "\n"
						+ "Date                : 10-05-2010\n"
						+ "Release No          : 1.0\n"
						+ "Author              : Satish Kumar\n"
						+ "Description         : This software is made to generate indic language .pdf files. \nIt uses the Open Office SDK to create .odf first and \nfollowed by the .pdf document.\n");
		// after insertString the cursor is behind the inserted text, insert
		// more text
		// xText.insertString(xTextCursor, "\n \"Is that the mill?\" he asked.",
		// false);
	}

	/**
	 * This method is created to demonstrate different language font on the PDF.
	 * 
	 * @param langId
	 * @return
	 */
	public static String getText(String langId) {
		String data = "";
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			java.sql.Connection connection = java.sql.DriverManager
					.getConnection("jdbc:oracle:thin:@192.168.20.18:1527:cld1",
							"esiuser1", "esiuser1");
			java.sql.Statement stmt = connection.createStatement();
			java.sql.ResultSet rs = stmt
					.executeQuery("select TEXT from I18N where TITLE='"
							+ langId + "'");

			// String key = "";

			// Play around with text
			while (rs.next()) {
				data += rs.getString("TEXT");
			}
		} catch (java.lang.Exception e) {
			System.err.println(e.getMessage());
			e.printStackTrace();
		}
		return data;
	}

	/**
	 * This will create the XComponentContext object using the bootstrap.
	 * 
	 * @return - XComponentContext object.
	 */
	private XComponentContext getContext() {
		XComponentContext xContext = null;

		try {
			// get the remote office component context
			xContext = BootstrapSocketConnector.bootstrap(OOO_EXEC_FOLDER);
			if (xContext != null)
				System.out.println("Connected to a running office ...");
		} catch (java.lang.Exception e) {
			e.printStackTrace(System.err);
			System.exit(1);
		}
		return xContext;
	}

	/**
	 * This method is responsible to store the component to PDF format.
	 * 
	 * @param xComponent -
	 *            component to store
	 * @param storeUrl -
	 *            url to store the component
	 * @throws Exception -
	 *             checked exception
	 */
	private void storePDFComponent(XComponent xComponent, String storeUrl)
			throws Exception {

		XStorable xStorable = (XStorable) UnoRuntime.queryInterface(
				XStorable.class, xComponent);
		PropertyValue[] storeProps = new PropertyValue[1];
		storeProps[0] = new PropertyValue();
		storeProps[0].Name = "FilterName";
		storeProps[0].Value = "writer_pdf_Export";

		System.out.println("... store  to \"" + storeUrl + "\".");
		xStorable.storeToURL(storeUrl, storeProps);
	}

	/**
	 * Set the XDocument proporties. Additional features can be added in future.
	 * 
	 * @return - property list.
	 */
	private PropertyValue[] setDocumentProperties() {
		ArrayList props = new ArrayList();
		PropertyValue propertyValue;

		propertyValue = new PropertyValue();
		propertyValue.Name = "Hidden";
		propertyValue.Value = new Boolean(true);

		props.add(propertyValue);

		PropertyValue[] properties = new PropertyValue[props.size()];
		props.toArray(properties);

		return properties;
	}

	/**
	 * Set the XDocument proporties. Additional features can be added in future.
	 * 
	 * @return - property list.
	 */
	private void setTableProperties(XTextTable xTextTable) {
		try {
			XPropertySet xPS = (XPropertySet) UnoRuntime.queryInterface(
					XPropertySet.class, xTextTable);

			xPS.setPropertyValue("HoriOrient", new Integer(
					com.sun.star.text.HoriOrientation.LEFT));

			xPS.setPropertyValue("TopMargin", new Integer(20));

			Integer iiWidth = (Integer) xPS.getPropertyValue("Width");

			System.out.println("Width : " + iiWidth + "|"
					+ xPS.getPropertyValue("IsWidthRelative"));
			xPS.setPropertyValue("BackTransparent", new Boolean(true));

			// Get table Width and TableColumnRelativeSum properties values
			int iWidth = ((Integer) xPS.getPropertyValue("Width")).intValue();
			short sTableColumnRelativeSum = ((Short) xPS
					.getPropertyValue("TableColumnRelativeSum")).shortValue();

			// Calculate conversion ration
			double dRatio = (double) sTableColumnRelativeSum / (double) iWidth;

			// Convert our 1 mm (100) to unknown ( relative ) units
			double dRelativeWidth1 = (double) 90000 * dRatio;

			// Get table column separators
			Object xObj = xPS.getPropertyValue("TableColumnSeparators");

			TableColumnSeparator[] xSeparators = (TableColumnSeparator[]) UnoRuntime
					.queryInterface(TableColumnSeparator[].class, xObj);

			// Last table column separator position
			double dPosition1 = sTableColumnRelativeSum - dRelativeWidth1;
			// Set set new position for all column separators
			xSeparators[0].Position = (short) Math.ceil(dPosition1);

			// Set TableColumnSeparators back!.
			xPS.setPropertyValue("TableColumnSeparators", xSeparators);

			// create description for blue line, width 10
			BorderLine theLine = new BorderLine();
			theLine.Color = 0xFFCFFF;
			theLine.OuterLineWidth = 1;
			theLine.InnerLineWidth = 0;
			theLine.LineDistance = 3;
			// apply line description to all border lines and make them valid
			TableBorder bord = new TableBorder();
			bord.VerticalLine = bord.HorizontalLine = bord.LeftLine = bord.RightLine = bord.TopLine = bord.BottomLine = theLine;
			bord.IsVerticalLineValid = bord.IsHorizontalLineValid = bord.IsLeftLineValid = bord.IsRightLineValid = bord.IsTopLineValid = bord.IsBottomLineValid = true;

			xPS.setPropertyValue("TableBorder", bord);

		} catch (UnknownPropertyException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (PropertyVetoException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * This method is responsible to create a table object and add to the
	 * provided document.
	 * 
	 * @param xTextDocument -
	 *            the text document to add the table object.
	 * @param xWriterComponent -
	 *            the writer component for the document.
	 * @return - table object.
	 */
	private XTextTable CreateTable(XTextDocument xTextDocument,
			XComponent xWriterComponent) {
		// insert TextTable and get cell text, then manipulate text in cell
		XTextTable xTextTable = null;
		try {
			XText xText = xTextDocument.getText();

			// manipulateText(xText);

			// get internal service factory of the document
			XMultiServiceFactory xWriterFactory = (XMultiServiceFactory) UnoRuntime
					.queryInterface(XMultiServiceFactory.class,
							xWriterComponent);

			xTextTable = (XTextTable) UnoRuntime.queryInterface(
					XTextTable.class, xWriterFactory
							.createInstance("com.sun.star.text.TextTable"));

			xTextTable.initialize(2, 2);

			XTextContent xTextContentTable = (XTextContent) UnoRuntime
					.queryInterface(XTextContent.class, xTextTable);

			xText.insertTextContent(xText.getEnd(), xTextContentTable, false);
		} catch (IllegalArgumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (java.lang.Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return xTextTable;
	}

	/**
	 * This method demonstrate the data arrangement in a table. This method also
	 * demonstrate how to work with the tables.
	 * 
	 * @param xCellRange
	 * @throws com.sun.star.uno.Exception
	 */
	private void CreateTableRow(XCellRange xCellRange, String data)
			throws com.sun.star.uno.Exception {

		XCell xCell = null;
		XText xCellText = null;
		XPropertySet xPS = null;
		XTextCursor xCellCursor = null;
		XPropertySet xCellCursorProps = null;

		xCell = xCellRange.getCellByPosition(0, 1);
		xCellText = (XText) UnoRuntime.queryInterface(XText.class, xCell);
		xPS = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class,
				xCell);
		xPS.setPropertyValue("VertOrient", new Integer(
				com.sun.star.text.VertOrientation.TOP));

		// create a text cursor from the cells XText interface
		xCellCursor = xCellText.createTextCursor();

		// Get the property set of the cell's TextCursor
		xCellCursorProps = (XPropertySet) UnoRuntime.queryInterface(
				XPropertySet.class, xCellCursor);

		// Set the colour of the text to white
		xCellCursorProps.setPropertyValue("CharWeight", new Float(
				com.sun.star.awt.FontWeight.BOLD));
		xCellCursorProps.setPropertyValue("CharFontName", new String(
				"Helvetica"));
		xCellCursorProps.setPropertyValue("CharHeight", new Float(10));
		xCellText.setString("Language Name");

		// enter column titles and a cell value - 2st ":"
		xCell = xCellRange.getCellByPosition(1, 1);
		xCellText = (XText) UnoRuntime.queryInterface(XText.class, xCell);
		xPS = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class,
				xCell);
		xPS.setPropertyValue("VertOrient", new Integer(
				com.sun.star.text.VertOrientation.TOP));
		// create a text cursor from the cells XText interface
		xCellCursor = xCellText.createTextCursor();

		// Get the property set of the cell's TextCursor
		xCellCursorProps = (XPropertySet) UnoRuntime.queryInterface(
				XPropertySet.class, xCellCursor);

		// Set the colour of the text to white
		// xCellCursorProps.setPropertyValue("CharWeight", new Float(
		// com.sun.star.awt.FontWeight.BOLD));
		xCellCursorProps.setPropertyValue("CharFontName", new String(
				"Helvetica"));
		xCellCursorProps.setPropertyValue("CharHeight", new Float(10));
		xCellText.setString(data);
	}

	private void fillFirstRow(XCellRange xCellRange){
		XCell xCell = null;
		XText xCellText = null;

		// Filling the first row to dummy value
		// enter column titles and a cell value
		try {
			xCell = xCellRange.getCellByPosition(0, 0);
			xCellText = (XText) UnoRuntime.queryInterface(XText.class, xCell);
			xCellText.setString("Language");

			// enter column titles and a cell value
			xCell = xCellRange.getCellByPosition(1, 0);
			xCellText = (XText) UnoRuntime.queryInterface(XText.class, xCell);
			xCellText.setString("Text");

		} catch (IndexOutOfBoundsException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	
	

}
