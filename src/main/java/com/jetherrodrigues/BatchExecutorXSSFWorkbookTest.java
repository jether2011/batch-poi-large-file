package com.jetherrodrigues;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * @author JETHER ROIS
 * 
 */
public class BatchExecutorXSSFWorkbookTest {

	private static String INPUTS_SHEET;
	private static final String DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";
	private static Workbook _wb;
	private static SXSSFWorkbook _sxssfWorkbook;

	public static void main(String[] args) {
		System.out.println("Start test: " + new SimpleDateFormat(DATE_PATTERN).format(new Date()));

		if (args.length != 2) {
			throw new IllegalArgumentException(
					"Missing arguments. Is need send two args: 1.the path file and 2.the sheet name.");
		}

		INPUTS_SHEET = args[1].toString().trim();

		File _file = new File(args[0].toString().trim());
		System.out.println("File created and loaded [" + _file.getAbsolutePath() + "]");

		OPCPackage _opc;
		try {
			System.out.println("Starting try open with StreamingReader.....");

			_opc = OPCPackage.open(_file, PackageAccess.READ_WRITE);
			_wb = new XSSFWorkbook(_opc);
			//_sxssfWorkbook = new SXSSFWorkbook((XSSFWorkbook) _wb, 500);
			
			System.out.println("XSSFWorkbook is ok [" + _wb.toString() + "]");

			Sheet _sheet = _wb.getSheet(INPUTS_SHEET);

			System.out.println("Sheet is ok [" + _sheet.getSheetName() + "] and starting loop to look inside sheet....");
			
//			Row row;
//			Iterator<Row> rows = _sheet.iterator();
//			
//			Cell cell;
//			Iterator<Cell> cells;
//			
//			while (rows.hasNext()) {
//				row = rows.next();
//				cells = row.iterator();
//				while (cells.hasNext()) {
//					cell = (Cell) cells.next();
//					System.out.println(cell);
//				}
//			}
			
			System.out.println("End the proccess of batch: " + new SimpleDateFormat(DATE_PATTERN).format(new Date()));

		} catch (Exception e) {
			System.out.println(e.getMessage());
		} finally {

		}
	}

}
