package com.example.demospringbootpoixlsx;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetPr;
import org.springframework.boot.test.context.SpringBootTest;

import java.awt.Color;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Base64;
import java.util.List;

@SpringBootTest
class DemoSpringbootPoiXlsxApplicationTests {

	@Test
	void contextLoads() {
	}

	private void changeCellBackgroundColor(Cell cell) {
		CellStyle cellStyle = cell.getCellStyle();
		if(cellStyle == null) {
			cellStyle = cell.getSheet().getWorkbook().createCellStyle();
		}
		cellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(cellStyle);
	}

	// source: https://www.iditect.com/faq/java/auto-size-height-for-rows-in-apache-poi.html
	private void resizeRowHeight(Row row) {
		int rowCount = row.getSheet().getPhysicalNumberOfRows();
		for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
			Row currentRow = row.getSheet().getRow(rowIndex);
			if (currentRow != null) {
				short height = 0;
				int cellCount = currentRow.getLastCellNum();
				for (int cellIndex = 0; cellIndex < cellCount; cellIndex++) {
					Cell cell = currentRow.getCell(cellIndex);
					if (cell != null) {
						int cellLength = cell.toString().length();
						height = (short) Math.max(height, (short) (cellLength * 256)); // Some factor for scaling
					}
				}
				currentRow.setHeight(height);
			}
		}
	}

	// Source https://stackoverflow.com/questions/13930668/add-border-to-merged-cells-in-excel-apache-poi-java
	// impact all merged regions in the sheet
	private void setBordersToMergedCells(Sheet sheet) {
		List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
		for (CellRangeAddress rangeAddress : mergedRegions) {
			RegionUtil.setBorderTop(BorderStyle.THICK, rangeAddress, sheet);
			RegionUtil.setBorderLeft(BorderStyle.THICK, rangeAddress, sheet);
			RegionUtil.setBorderRight(BorderStyle.THICK, rangeAddress, sheet);
			RegionUtil.setBorderBottom(BorderStyle.THICK, rangeAddress, sheet);
		}
	}

	@Test
	void createExcel() throws IOException {
		// Create a new workbook
		XSSFWorkbook workbook = new XSSFWorkbook();

		// Create a new sheet for testing color
		XSSFSheet colorSheet = workbook.createSheet("Color");

		// How to set color (first way)
		XSSFColor colorRed = new XSSFColor(new Color(0xFF0000), null);

		// How to set a cell background color
		XSSFCellStyle cellStyleRedBackground = workbook.createCellStyle();
		cellStyleRedBackground.setFillForegroundColor(colorRed);
		cellStyleRedBackground.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		// Try to set tab color (first way, does not work)
		colorSheet.setTabColor(colorRed);

		// Try to set tab color (second way, does not work)
		CTSheetPr pr = colorSheet.getCTWorksheet().addNewSheetPr();
		CTColor tabColor = pr.addNewTabColor();
		tabColor.setRgb(new byte[] { (byte) 0, (byte) 0, (byte) 255 }); // Blue
		pr.setTabColor(tabColor); 		//does not work

		// Create a first row on color sheet
		XSSFRow firstRowOnColorSheet = colorSheet.createRow(0);

		// How to set color (second way) and apply to cell style
		XSSFColor myColor = new XSSFColor(new Color(242, 220, 219), null); // #f2dcdb
		XSSFCellStyle cellStyleWithMyColorBackground = workbook.createCellStyle();
		cellStyleWithMyColorBackground.setFillForegroundColor(myColor);
		cellStyleWithMyColorBackground.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		// Create a new cell in the first row
		XSSFCell cellA1OnColorSheet = firstRowOnColorSheet.createCell(0);
		cellA1OnColorSheet.setCellStyle(cellStyleWithMyColorBackground);

		// Set value to cell A1 on color sheet
		cellA1OnColorSheet.setCellValue("Hello World");

		byte[] greenWithByteColor = new byte[]{
				(byte) 0,	// Red
				(byte) 255, 	// Green
				(byte) 0	// Blue
		};
		XSSFColor colorGreen = new XSSFColor(greenWithByteColor, null);

		XSSFCellStyle cellStyleWithGreenColor = workbook.createCellStyle();
		cellStyleWithGreenColor.setFillForegroundColor(colorGreen);
		cellStyleWithGreenColor.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		//XSSFCell cellA2OnColorSheet = row.createCell(1);
		XSSFCell cellA3OnColorSheet = firstRowOnColorSheet.createCell(2);
		//changeCellBackgroundColor(cellA3OnColorSheet);
		cellA3OnColorSheet.setCellStyle(cellStyleWithGreenColor);

		// Add a new sheet to test image
		XSSFSheet imageSheet = workbook.createSheet("Image");

		XSSFRow firstRowOnImageSheet = imageSheet.createRow(0);
		XSSFCell cellA1OnImageSheet = firstRowOnImageSheet.createCell(0);
		XSSFCell cellA2OnImageSheet = firstRowOnImageSheet.createCell(1);

		cellA1OnImageSheet.setCellValue("Logo");

		//Using https://base64.guru/converter/encode/image/png to convert image (URL) to base64
		// Remote url: https://base64.pi7.org/static/base4_Image.png
		String logoBase64 = "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";
		byte[] inputImageBytes1 = Base64.getDecoder().decode(logoBase64);

		// Supported Image: PNG, JPG, ... (look for Workbook.PICTURE_TYPE_...)
		int inputImagePictureID1 = workbook.addPicture(inputImageBytes1, Workbook.PICTURE_TYPE_PNG);

		// Create Drawing Container
		XSSFDrawing drawing = (XSSFDrawing) imageSheet.createDrawingPatriarch();

		// Create an anchor that is attached to the worksheet
		XSSFClientAnchor drawingAnchor = new XSSFClientAnchor();
		drawingAnchor.setRow1(0);
		drawingAnchor.setRow2(1);
		drawingAnchor.setCol1(1);
		drawingAnchor.setCol2(2);
		drawing.createPicture(drawingAnchor, inputImagePictureID1);

		// How to resize the row height and column width to fit the image
		// make sure the cells are wide enough for the pictures we’ve added by using autoSizeColumn:
		for (int i = 0; i < 3; i++) {
			imageSheet.autoSizeColumn(i);
		}
		resizeRowHeight(firstRowOnImageSheet);

		// Add new sheet to test font
		XSSFSheet fontSheet = workbook.createSheet("Font");

		// Source: https://www.baeldung.com/apache-poi-change-cell-font
		// Do not use a cell’s getCellStyle method that will always return a non-null value -> confusing and will modify all cell styles
		// A CellStyle is scoped to a workbook and is not specific to a sheet or a row or a cell
		CellStyle warningCellStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontName("Courier New");
		font.setBold(true);
		font.setFontHeight((short) 400);
		font.setUnderline(Font.U_SINGLE);
		// U_NONE	Text without underline
		// U_SINGLE	Single underline text where only the word is underlined
		// U_SINGLE_ACCOUNTING	Single underline text where almost the entire cell width is underlined
		// U_DOUBLE	Double underline text where only the word is underlined
		// U_DOUBLE_ACCOUNTING	Double underline text where almost the entire cell width is underlined
		font.setColor(IndexedColors.RED.getIndex());
		warningCellStyle.setFont(font);
		warningCellStyle.setAlignment(HorizontalAlignment.CENTER);
		warningCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		// Create a row in fontSheet and put some cells in it. Rows are 0 based.
		XSSFRow firstRowOnFontSheet = fontSheet.createRow(0);
		XSSFCell cellA1OnFontSheet = firstRowOnFontSheet.createCell(0);
		cellA1OnFontSheet.setCellValue("Courier New Bold Red Underline Warning");
		cellA1OnFontSheet.setCellStyle(warningCellStyle);
		firstRowOnFontSheet.setHeightInPoints((short) 40);

		XSSFRow thirdRowOnFontSheet = fontSheet.createRow(2);
		XSSFCell cellA3OnFontSheet = thirdRowOnFontSheet.createCell(0);
		cellA3OnFontSheet.setCellValue("Another text");

		// Add new sheet to test cell border
		XSSFSheet borderSheet = workbook.createSheet("Border");
		XSSFRow secondRowOnBorderSheet = borderSheet.createRow(1);
		XSSFCell cellB2OnBorderSheet = secondRowOnBorderSheet.createCell(1);
		cellB2OnBorderSheet.setCellValue("Cell");

		// Create a border for the cell
		CellStyle borderCellStyle = workbook.createCellStyle();
		borderCellStyle.setBorderBottom(BorderStyle.THICK);
		borderCellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
		borderCellStyle.setBorderLeft(BorderStyle.THICK);
		borderCellStyle.setLeftBorderColor(IndexedColors.BLUE.getIndex());
		borderCellStyle.setBorderRight(BorderStyle.THICK);
		borderCellStyle.setRightBorderColor(IndexedColors.RED.getIndex());
		borderCellStyle.setBorderTop(BorderStyle.THICK);
		borderCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
		cellB2OnBorderSheet.setCellStyle(borderCellStyle);

		// Try with merging cells
		// source: https://www.baeldung.com/java-apache-poi-merge-cells
		XSSFRow forthRowOnBorderSheet = borderSheet.createRow(3);
		XSSFCell cellB4OnBorderSheet = forthRowOnBorderSheet.createCell(1);
		XSSFCell cellC4OnBorderSheet = forthRowOnBorderSheet.createCell(2);
		//cellB4OnBorderSheet.setCellValue("Cell merged !!");
		int firstRow = 3;
		int lastRow = 3;
		int firstCol = 1;
		int lastCol = 2;
		CellRangeAddress rangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
		fontSheet.addMergedRegion(rangeAddress); //first way
		cellB4OnBorderSheet.setCellValue("Cell merged !!");
		//cellB4OnBorderSheet.setCellStyle(borderCellStyle); // do not work, right border is not displayed
		//fontSheet.addMergedRegion(CellRangeAddress.valueOf("B4:C4")); //second way !!Carefull
		//setBordersToMergedCells(fontSheet); //impact all regions for the sheet

		// Limitation: cannot use color to region border
		RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, borderSheet);
		RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, borderSheet);
		RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, borderSheet);
		RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, borderSheet);

		// Create a new sheet to test formula
		// Help: https://poi.apache.org/components/spreadsheet/eval.html
		// Source: https://blog.fileformat.com/spreadsheet/work-with-excel-formulas-in-java-applications-with-apache-poi-library/
		XSSFSheet formulaSheet = workbook.createSheet("Formula");
		XSSFRow formulaRow = formulaSheet.createRow(1);
		XSSFCell formulaCell = formulaRow.createCell(1);
		formulaCell.setCellValue("A = ");
		formulaCell = formulaRow.createCell(2);
		formulaCell.setCellValue(2);
		formulaRow = formulaSheet.createRow(2);
		formulaCell = formulaRow.createCell(1);
		formulaCell.setCellValue("B = ");
		formulaCell = formulaRow.createCell(2);
		formulaCell.setCellValue(4);
		formulaRow = formulaSheet.createRow(3);
		formulaCell = formulaRow.createCell(1);
		formulaCell.setCellValue("Total = ");
		formulaCell = formulaRow.createCell(2);
		// Create SUM formula
		formulaCell.setCellFormula("SUM(C2:C3)");
		formulaCell = formulaRow.createCell(3);
		formulaCell.setCellValue("SUM(C2:C3)");
		formulaRow = formulaSheet.createRow(4);
		formulaCell = formulaRow.createCell(1);
		formulaCell.setCellValue("POWER =");
		formulaCell=formulaRow.createCell(2);
		// Create POWER formula
		formulaCell.setCellFormula("POWER(C2,C3)");
		formulaCell = formulaRow.createCell(3);
		formulaCell.setCellValue("POWER(C2,C3)");
		formulaRow = formulaSheet.createRow(5);
		formulaCell = formulaRow.createCell(1);
		formulaCell.setCellValue("MAX = ");
		formulaCell = formulaRow.createCell(2);
		// Create MAX formula
		formulaCell.setCellFormula("MAX(C2,C3)");
		formulaCell = formulaRow.createCell(3);
		formulaCell.setCellValue("MAX(C2,C3)");
		formulaRow = formulaSheet.createRow(6);
		formulaCell = formulaRow.createCell(1);
		formulaCell.setCellValue("FACT = ");
		formulaCell = formulaRow.createCell(2);
		// Create FACT formula
		formulaCell.setCellFormula("FACT(C3)");
		formulaCell = formulaRow.createCell(3);
		formulaCell.setCellValue("FACT(C3)");
		formulaRow = formulaSheet.createRow(7);
		formulaCell = formulaRow.createCell(1);
		formulaCell.setCellValue("SQRT = ");
		formulaCell = formulaRow.createCell(2);
		// Create SQRT formula
		formulaCell.setCellFormula("SQRT(C5)");
		formulaCell = formulaRow.createCell(3);
		formulaCell.setCellValue("SQRT(C5)");

		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		evaluator.evaluateAll();

		// Create a new sheet to test hyperlink
		XSSFSheet hyperlinkSheet = workbook.createSheet("Hyperlink");
		XSSFRow hyperlinkRow = hyperlinkSheet.createRow(0);
		XSSFCell hyperlinkCell = hyperlinkRow.createCell(0);
		hyperlinkCell.setCellValue("Click here to go to Google");
		XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
		link.setAddress("http://www.google.com");
		hyperlinkCell.setHyperlink(link);

		// Retrieve cell value from another sheet
		CellAddress cellAddress = new CellAddress("C4");
		Row rowFormulaC4 = formulaSheet.getRow(cellAddress.getRow());
		Cell cellFormulaC4 = rowFormulaC4.getCell(cellAddress.getColumn());
		XSSFRow hyperlinkRow2 = hyperlinkSheet.createRow(1);
		XSSFCell hyperlinkCell2Label = hyperlinkRow2.createCell(0);
		hyperlinkCell2Label.setCellValue("Sum");
		XSSFCell hyperlinkCell2 = hyperlinkRow2.createCell(1);
		if (cellFormulaC4.getCellType() == CellType.FORMULA) {
			switch (evaluator.evaluateFormulaCell(cellFormulaC4)) {
				case BOOLEAN:
					hyperlinkCell2.setCellValue(cellFormulaC4.getBooleanCellValue());
					break;
				case NUMERIC:
					hyperlinkCell2.setCellValue(cellFormulaC4.getNumericCellValue());
					break;
				case STRING:
					hyperlinkCell2.setCellValue(cellFormulaC4.getStringCellValue());
					break;
			}
		} else {
			switch (cellFormulaC4.getCellType()) {
				case BOOLEAN:
					hyperlinkCell2.setCellValue(cellFormulaC4.getBooleanCellValue());
					break;
				case NUMERIC:
					hyperlinkCell2.setCellValue(cellFormulaC4.getNumericCellValue());
					break;
				case STRING:
					hyperlinkCell2.setCellValue(cellFormulaC4.getStringCellValue());
					break;
			}
		}

		// Create a new sheet to test later modification
		XSSFSheet modificationSheet = workbook.createSheet("Modification");
		XSSFRow modificationRow = modificationSheet.createRow(0);
		XSSFCell modificationLabelCell = modificationRow.createCell(0);
		XSSFCell modificationCell = modificationRow.createCell(1);
		modificationLabelCell.setCellValue("The following cell will be modified later");
		modificationCell.setCellValue("Not modified");

		// Write the output to a file
		try (FileOutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
			workbook.write(fileOut);
		} catch (IOException e) {
			e.printStackTrace();
		}
		workbook.close();
	}

	@Test
	void modifyExcel() throws IOException {
		FileInputStream file = new FileInputStream(new File("workbook.xlsx"));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet modificationSheet = workbook.getSheet("Modification");
		XSSFRow modificationRow = modificationSheet.getRow(0);
		XSSFCell modificationCell = modificationRow.getCell(1);
		modificationCell.setCellValue("Modified");

		file.close();

		// Write the output to a file
		try (FileOutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
			workbook.write(fileOut);
		} catch (IOException e) {
			e.printStackTrace();
		}
		workbook.close();

	}

}
