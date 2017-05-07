package br.ff.test;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.SheetUtil;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;
/**
 * 
 * @author JETHER ROIS
 * 
 * http://www.consulting-bolte.de/index.php/apache-poi/148-reading-big-excel-files-with-poi
 * 
 * https://myjeeva.com/read-excel-through-java-using-xssf-and-sax-apache-poi.html
 * 
 * https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xssf/eventusermodel/XLSX2CSV.java
 * 
 * http://www.massapi.com/class/xs/XSSFReader.html
 * 
 * http://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api
 *
 */
public class FFTest {
	private static OPCPackage _opc;
	private static final String INPUTS_SHEET = "Inputs";
	
	public static void main(String[] args) {
		File _file = new File(args[0].toString().trim());
		System.out.println("File created [" + _file.getAbsolutePath() + "]");
		try {
			_opc = OPCPackage.open(_file, PackageAccess.READ_WRITE);
			System.out.println("OPCPackage opened the file [" + _opc.toString() + "]");

			XSSFReader _reader = new XSSFReader(_opc);
			SharedStringsTable _sst = _reader.getSharedStringsTable();	
			StylesTable styles = _reader.getStylesTable();
			
			XMLReader _parser = fetchSheetParser(_sst);
						
			XSSFReader.SheetIterator _sheetIteration = (XSSFReader.SheetIterator) _reader.getSheetsData();
			
			int index = 0;
	        while (_sheetIteration.hasNext()) {
	        	InputStream _sheetStream = _sheetIteration.next();	        	
	        	if (_sheetIteration.getSheetName().compareToIgnoreCase(INPUTS_SHEET) == 0) {
	        		InputSource _source = new InputSource(_sheetStream);
	        		System.out.println("InputSource._source [" + _source.toString() + "]");
				}	        	
//	            InputStream stream = _sheetIteration.next();
//	            String sheetName = _sheetIteration.getSheetName();
//	            System.out.println(sheetName + " [index=" + index + "]:");
//	            processSheet(styles, strings, new SheetToCSV(), stream);
//	            stream.close();
	            ++index;
	        }

			//XMLReader _parser = fetchSheetParser(_sst);
			//System.out.println("XMLReader parser finished [" + _parser.toString() + "]");
			
			// process the first sheet (Inputs to FastForecast)
//			InputStream _sheet = _reader.getSheetsData().next();
//			InputSource _sheetSource = new InputSource(_sheet);
//			_parser.parse(_sheetSource);
//			_sheet.close();
			
			SXSSFWorkbook _wb = new SXSSFWorkbook(100);
			_wb.createSheet("Inputs");
			

		} catch (IOException | OpenXML4JException | SAXException e) {
			System.out.println(e.getMessage());
		} finally {
//			try {
//				_opc.close();
//			} catch (IOException e) {
//				System.out.println("There is a error in close _opc object. [ " + e.getMessage() + " ]");
//			}
		}
	}

	public static XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader();
		ContentHandler handler = new SheetHandler(sst);
		parser.setContentHandler(handler);
		return parser;
	}

	/**
	 * See org.xml.sax.helpers.DefaultHandler javadocs
	 */
	private static class SheetHandler extends DefaultHandler {
		private final SharedStringsTable sst;
		private String lastContents;
		private boolean nextIsString;
		private boolean inlineStr;
		private final LruCache<Integer, String> lruCache = new LruCache<Integer, String>(50);

		@SuppressWarnings("serial")
		private static class LruCache<A, B> extends LinkedHashMap<A, B> {
			private final int maxEntries;

			public LruCache(final int maxEntries) {
				super(maxEntries + 1, 1.0f, true);
				this.maxEntries = maxEntries;
			}

			@Override
			protected boolean removeEldestEntry(final Map.Entry<A, B> eldest) {
				return super.size() > maxEntries;
			}
		}

		private SheetHandler(SharedStringsTable sst) {
			this.sst = sst;
		}

		@Override
		public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
			// c => cell
			if (name.equals("c")) {
				// Print the cell reference
				System.out.print(attributes.getValue("r") + " - ");
				// Figure out if the value is an index in the SST
				String cellType = attributes.getValue("t");
				nextIsString = cellType != null && cellType.equals("s");
				inlineStr = cellType != null && cellType.equals("inlineStr");
			}
			// Clear contents cache
			lastContents = "";
		}

		@Override
		public void endElement(String uri, String localName, String name) throws SAXException {
			// Process the last contents as required.
			// Do now, as characters() may be called more than once
			if (nextIsString) {
				Integer idx = Integer.valueOf(lastContents);
				lastContents = lruCache.get(idx);
				if (lastContents == null && !lruCache.containsKey(idx)) {
					lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
					lruCache.put(idx, lastContents);
				}
				nextIsString = false;
			}

			// v => contents of a cell
			// Output after we've seen the string contents
			if (name.equals("v") || (inlineStr && name.equals("c"))) {
				System.out.println(lastContents);
			}
		}

		@Override
		public void characters(char[] ch, int start, int length) throws SAXException { // NOSONAR
			lastContents += new String(ch, start, length);
		}
	}
	
}
