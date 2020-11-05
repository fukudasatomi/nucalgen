package jp.satomichan.nucalgen;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFName;

public class Nucalgen {

	public static void main(String[] args) {
		//コマンドライン・オプション読み込み
		Options options = new Options();
		options.addOption(Option.builder("s").required().hasArg().longOpt("std-food-comp").build());
		options.addOption(Option.builder("c").required().hasArg().longOpt("columns").build());
		options.addOption(Option.builder("o").required().hasArg().longOpt("output").build());
		options.addOption(Option.builder("l").required().hasArg().longOpt("lines").build());
		options.addOption(Option.builder("r").longOpt("use-cache-std-food-comp").build());
		options.addOption(Option.builder("pfcbalance").build());

		try {

		CommandLineParser parser = new DefaultParser();
		CommandLine cmd = parser.parse(options, args);


		final String moeStdFoodCompTableFileName = cmd.getOptionValue("s");
		final String columnsXmlFileName = cmd.getOptionValue("c");
		final String outputXlsxFileName = cmd.getOptionValue("o");
		final int lines = Integer.parseInt(cmd.getOptionValue("l"));


			//コンフィグ読み込み
			XMLConfiguration config = new XMLConfiguration(columnsXmlFileName);
			NutritionColumnHolder nc = new NutritionColumnHolder(config);
			//System.out.println(nc);

			//Book生成
			Workbook outputWorkbook = WorkbookFactory.create(new FileInputStream(moeStdFoodCompTableFileName));
			//Workbook outputWorkbook = new XSSFWorkbook();

			if(cmd.hasOption("use-cache-std-food-comp") == false) {
				//「本表」変換
				Sheet mainSheet = outputWorkbook.getSheet("本表");
				int rowCount = 0;
				for (Row row : mainSheet) {
					rowCount++;
					if(rowCount < 8) {continue;}

					for (Cell cell : row) {
						String cellString = cell.toString();

						cellString = cellString.replaceAll("\\(", "");
						cellString = cellString.replaceAll("\\)", "");
						cellString = cellString.replaceAll("-", "0");
						cellString = cellString.replaceAll("Tr", "0");


						if(cellString.matches("^[\\d\\.]+$")) {
							cell.setCellValue(Double.parseDouble(cellString));
							CellStyle aCellStyle = cell.getCellStyle();
							aCellStyle.setDataFormat((short) 0);
							cell.setCellStyle(aCellStyle);
						}


						//System.out.print(cell.toString());
						//System.out.print(" , ");
					}
					//System.out.println();
				}

				mainSheet.getRow(5).getCell(4).setCellValue("廃棄率");

			}


			//「栄養価計算」シート生成
			Sheet calcSheet = outputWorkbook.createSheet("栄養価計算");
			outputWorkbook.setSheetOrder("栄養価計算", 0);
			//outputWorkbook.setActiveSheet(0);
			calcSheet.setColumnWidth(2, 10240);
			calcSheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));

			CellStylePool csPool = new CellStylePool(outputWorkbook);

			//「タイトル」行
			Row titleRow = calcSheet.createRow(1);
			titleRow.createCell(1).setCellValue("食品番号");
			titleRow.createCell(2).setCellValue("食品名");
			titleRow.createCell(3).setCellValue("摂取量");
			int colIndex = 4;
			for(NutritionColumn aColumn : nc.getNutritionColumnList()) {
				titleRow.createCell(colIndex).setCellValue(aColumn.getDispName());
				colIndex++;
			}

			//「単位」行
			Row unitRow = calcSheet.createRow(2);
			unitRow.createCell(2).setCellValue("単位");
			unitRow.createCell(3).setCellValue("g");
			colIndex = 4;
			for(NutritionColumn aColumn : nc.getNutritionColumnList()) {
				unitRow.createCell(colIndex).setCellValue(aColumn.getUnit());
				colIndex++;
			}

			//「栄養計算」行
			int rowIndex = 3;
			for(int i = rowIndex; i < lines + 3; i++,rowIndex++) {
				Row thisRow = calcSheet.createRow(rowIndex);

				thisRow.createCell(1).setCellStyle(csPool.getCellStyle("00000"));
				thisRow.createCell(2).setCellFormula("IFERROR(VLOOKUP(B" + (rowIndex + 1) + ",本表!$B$9:$BO$2199,3,FALSE),\"\")");

				colIndex = 4;
				for(NutritionColumn aColumn : nc.getNutritionColumnList()) {
					Cell thisCell = thisRow.createCell(colIndex);
					thisCell.setCellStyle(csPool.getCellStyle(aColumn.getFormat()));

					String div100 = aColumn.isUseRawValue() ? "" :  "/ 100 * $D" + (rowIndex + 1);

					thisCell.setCellFormula("IFERROR(VLOOKUP($B" + (rowIndex + 1) + ",本表!$B$9:$BO$2199,MATCH(\"" + aColumn.getName() + "\",本表!$B$6:$BO$6,0),FALSE) " + div100 + ",\"\")");
					colIndex++;
				}

			}

			//「合計」行
			Row sumRow = calcSheet.createRow(rowIndex);
			sumRow.createCell(1).setCellValue("合計");
			calcSheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
			colIndex = 4;
			for(NutritionColumn aColumn : nc.getNutritionColumnList()) {
				Cell thisCell = sumRow.createCell(colIndex);
				String sumArea = new CellReference(calcSheet.getSheetName(), 3, colIndex, true, true).formatAsString() + ":" + new CellReference(calcSheet.getSheetName(), rowIndex -1, colIndex, true, true).formatAsString();

				//名前付き範囲（alias あれば設定）
				if(aColumn.getAlias().length() > 0) {
					XSSFName namedRangeSum = (XSSFName) outputWorkbook.createName();
					namedRangeSum.setNameName("SUM_" + aColumn.getAlias());
					namedRangeSum.setRefersToFormula(new CellReference(calcSheet.getSheetName(), rowIndex, colIndex, true, true).formatAsString());
					XSSFName namedRangeArea = (XSSFName) outputWorkbook.createName();
					namedRangeArea.setNameName("AREA_" + aColumn.getAlias());
					namedRangeArea.setRefersToFormula(sumArea);

				}

				thisCell.setCellStyle(csPool.getCellStyle(aColumn.getFormat()));
				if(aColumn.isUseSum()) {
					thisCell.setCellFormula("SUM(" + sumArea + ")");
				}
				colIndex++;
			}


			//PFCバランス出力
			if(cmd.hasOption("pfcbalance")) {
				rowIndex += 3;
				Row pfbBalanceRow1 = calcSheet.createRow(rowIndex);
				pfbBalanceRow1.createCell(1).setCellValue("PFCバランス (%)");
				calcSheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 2));
				pfbBalanceRow1.createCell(3).setCellValue("P");
				pfbBalanceRow1.createCell(4).setCellValue("F");
				pfbBalanceRow1.createCell(5).setCellValue("C");

				rowIndex++;
				Row pfbBalanceRow2 = calcSheet.createRow(rowIndex);
				Cell pCell = pfbBalanceRow2.createCell(3);
				pCell.setCellStyle(csPool.getCellStyle("0"));
				pCell.setCellFormula("SUM_P*4*100/(SUM_P*4+SUM_F*9+SUM_C*4)");
				Cell fCell = pfbBalanceRow2.createCell(4);
				fCell.setCellStyle(csPool.getCellStyle("0"));
				fCell.setCellFormula("SUM_F*9*100/(SUM_P*4+SUM_F*9+SUM_C*4)");
				Cell cCell = pfbBalanceRow2.createCell(5);
				cCell.setCellStyle(csPool.getCellStyle("0"));
				cCell.setCellFormula("SUM_C*4*100/(SUM_P*4+SUM_F*9+SUM_C*4)");
			}


			//ブック出力
            FileOutputStream outputXlsxFile = new FileOutputStream(outputXlsxFileName);
			outputWorkbook.setActiveSheet(0);
			calcSheet.setForceFormulaRecalculation(true);
			outputWorkbook.setSelectedTab(0);
            outputWorkbook.write(outputXlsxFile);


            outputWorkbook.close();
		} catch (Exception e) {
			// TODO 自動生成された catch ブロック
			e.printStackTrace();
		}
	}





}
