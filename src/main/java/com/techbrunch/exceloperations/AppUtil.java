package com.techbrunch.exceloperations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URL;
import java.net.URLDecoder;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class AppUtil {

	private void writeToExcel(final JsonObject jsonObjFromFile) throws IOException {
		int rowNum = 0;
		int counter = 0;
		final JsonArray columnHeaders = jsonObjFromFile.getAsJsonArray(AppConstants.COLUMN_HEADERS);
		final JsonArray dataRows = jsonObjFromFile.getAsJsonArray(AppConstants.DATA_ROWS_LABEL);
		final HSSFWorkbook workBook = new HSSFWorkbook();
		final HSSFSheet workSheet = workBook.createSheet(AppConstants.SHEET_NAME);
		for (final JsonElement dataElement : dataRows) {
			if (rowNum == 0) {
				Row row = workSheet.createRow(rowNum++);
				int colNum = 0;
				for (final JsonElement headerElement : columnHeaders) {
					row.createCell(colNum++).setCellValue(headerElement.getAsString());
				}
			}
			Row row = workSheet.createRow(rowNum++);
			int colNum = 0;
			row.createCell(colNum++).setCellValue(++counter);
			row.createCell(colNum++)
					.setCellValue(dataElement.getAsJsonObject().get(AppConstants.EMP_NAME).getAsString());
			row.createCell(colNum++)
					.setCellValue(dataElement.getAsJsonObject().get(AppConstants.EMP_AGE).getAsString());
			row.createCell(colNum++)
					.setCellValue(dataElement.getAsJsonObject().get(AppConstants.EMP_RATING).getAsString());
			row.createCell(colNum++)
					.setCellValue(dataElement.getAsJsonObject().get(AppConstants.EMP_SALARY).getAsString());
		}
		FileOutputStream outputStream = new FileOutputStream(AppConstants.PATH_TO_EXCEL_FILE);
		workBook.write(outputStream);
		workBook.close();
	}

	private JsonObject readFromJsonFile() throws FileNotFoundException, UnsupportedEncodingException {
		final URL urlToFile = getClass().getClassLoader().getResource(AppConstants.JSON_FILE_NAME);
		final FileReader fileReader = new FileReader(URLDecoder.decode(urlToFile.getFile(), AppConstants.UTF_8));
		final JsonElement element = new JsonParser().parse(fileReader);
		final JsonObject jsonObject = element.getAsJsonObject();
		return jsonObject;
	}

	private void readFromExcel(final String pathToExcel) throws IOException {
		final FileInputStream excelFile = new FileInputStream(new File(AppConstants.PATH_TO_EXCEL_FILE));
		final HSSFWorkbook workBook = new HSSFWorkbook(excelFile);
		final HSSFSheet workSheet = workBook.getSheetAt(0);
		Iterator<Row> rowIterator = workSheet.iterator();
		while (rowIterator.hasNext()) {
			final Row currentRow = rowIterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				System.out.print(currentCell.toString() + " ");
			}
			System.out.println();
		}
		workBook.close();
	}

	public void performAction(final int userResponse) throws IOException {
		switch (userResponse) {
			case 1: {
				// This case signifies that write to excel needs to be performed.
				final JsonObject jsonObjFromFile = readFromJsonFile();
				writeToExcel(jsonObjFromFile);
				break;
			}
			case 2: {
				// This case signifies that read from excel needs to be performed.
				readFromExcel(AppConstants.PATH_TO_EXCEL_FILE);
				break;
			}
		}
	}
}
