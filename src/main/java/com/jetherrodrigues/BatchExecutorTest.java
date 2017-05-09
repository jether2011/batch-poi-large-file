package com.jetherrodrigues;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.monitorjbl.xlsx.StreamingReader;

/**
 * 
 * @author JETHER ROIS
 * 
 *         https://github.com/monitorjbl/excel-streaming-reader
 */
public class BatchExecutorTest {

	private static String INPUTS_SHEET;
	private static final String DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";

	public static void main(String[] args) {
		System.out.println("Start test: " + new SimpleDateFormat(DATE_PATTERN).format(new Date()));

		if (args.length != 2) {
			throw new IllegalArgumentException(
					"Missing arguments. Is need send two args: 1.the path file and 2.the sheet name.");
		}

		INPUTS_SHEET = args[0].toString().trim();

		File _file = new File(args[1].toString().trim());
		System.out.println("File created and loaded [" + _file.getAbsolutePath() + "]");

		InputStream _is;
		try {
			System.out.println("Starting try open with StreamingReader.....");

			_is = new FileInputStream(_file);
			Workbook _wb = StreamingReader.builder().rowCacheSize(200).bufferSize(4096).open(_is);
			
			System.out.println("StreamingReader is ok [" + _wb.toString() + "]");

			Sheet _sheet = _wb.getSheet(INPUTS_SHEET);

			System.out.println("Sheet is ok [" + _sheet.getSheetName() + "] and starting loop to look inside sheet....");

			for (Row _r : _sheet) {
				for (Cell _c : _r) {
					System.out.println(_c.getStringCellValue());
				}
			}

			System.out.println("End the proccess of batch: " + new SimpleDateFormat(DATE_PATTERN).format(new Date()));

		} catch (Exception e) {
			System.out.println(e.getMessage());
		} finally {

		}
	}

}
