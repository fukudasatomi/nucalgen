package jp.satomichan.nucalgen;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.List;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.poi.ss.usermodel.Cell;
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
		options.addOption(Option.builder("pfcbalance").longOpt("with-pfc-balance").build());
		options.addOption(Option.builder("groupsum").longOpt("with-group-sum").build());
		options.addOption(Option.builder("bright").hasArg().longOpt("bright-colored-vegetables").build());

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

			//Book生成
			Workbook outputWorkbook = WorkbookFactory.create(new FileInputStream(moeStdFoodCompTableFileName));

			if(cmd.hasOption("use-cache-std-food-comp") == false) {
				//「本表」変換
				MoeStdFoodCompTable moe = new MoeStdFoodCompTable(cmd.getOptionValue("bright"));
				moe.convert(outputWorkbook);
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
				thisRow.createCell(2).setCellFormula("IFERROR(VLOOKUP(B" + (rowIndex + 1) + ",本表!$B$9:$BS$2199,3,FALSE),\"\")");

				colIndex = 4;
				for(NutritionColumn aColumn : nc.getNutritionColumnList()) {
					Cell thisCell = thisRow.createCell(colIndex);
					thisCell.setCellStyle(csPool.getCellStyle(aColumn.getFormat()));

					String div100 = aColumn.isUseRawValue() ? "" :  "/ 100 * $D" + (rowIndex + 1);

					thisCell.setCellFormula("IFERROR(VLOOKUP($B" + (rowIndex + 1) + ",本表!$B$9:$BS$2199,MATCH(\"" + aColumn.getName() + "\",本表!$B$6:$BS$6,0),FALSE) " + div100 + ",\"\")");
					colIndex++;
				}

			}


			//摂取量　名前付き範囲
			String intakeArea = new CellReference(calcSheet.getSheetName(), 3, 3, true, true).formatAsString() + ":" + new CellReference(calcSheet.getSheetName(), rowIndex -1, 3, true, true).formatAsString();
			XSSFName intakeNamedRangeArea = (XSSFName) outputWorkbook.createName();
			intakeNamedRangeArea.setNameName("AREA_INTAKE");
			intakeNamedRangeArea.setRefersToFormula(intakeArea);


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


			//「PFCバランス」出力
			if(cmd.hasOption("pfcbalance")) {
				rowIndex += 3;
				rowIndex = generatePfcBalance(calcSheet, csPool, rowIndex);
			}

			//「食品群別摂取量」出力
			if(cmd.hasOption("groupsum")) {
				rowIndex += 3;
				rowIndex = generateGroupSum(calcSheet, csPool, rowIndex);
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



	//PFCバランス
	private static int generatePfcBalance(Sheet calcSheet, CellStylePool csPool, int rowIndex) {
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

		return rowIndex;
	}



	 //群別摂取量
	private static int generateGroupSum(Sheet calcSheet, CellStylePool csPool, int rowIndex) {

		List<String> groupName = Arrays.asList("0", "穀類", "いも及びでん粉類", "砂糖及び甘味類", "豆類",
				"種実類", "野菜類　合計", "果実類", "きのこ類", "藻類", "魚介類", "肉類", "卵類", "乳類",
				"油脂類", "菓子類", "し好飲料類", "調味料及び香辛料類", "調理加工食品類");

		Row groupRow = calcSheet.createRow(rowIndex);
		groupRow.createCell(1).setCellValue("食品群");
		groupRow.createCell(3).setCellValue("摂取量(g)");
		rowIndex++;

		for(int i = 1; i <= 18; i++,rowIndex++) {
			Row thisRow = calcSheet.createRow(rowIndex);
			thisRow.createCell(1).setCellValue(i);
			thisRow.createCell(2).setCellValue(groupName.get(i));
			Cell cCell = thisRow.createCell(3);
			cCell.setCellStyle(csPool.getCellStyle("0"));
			cCell.setCellFormula("SUMIF(AREA_GROUP, " + i + ", AREA_INTAKE)");

			if(i == 6) {
				rowIndex++;
				thisRow = calcSheet.createRow(rowIndex);
				thisRow.createCell(2).setCellValue("うち　緑黄色野菜");
				Cell bcvCell = thisRow.createCell(3);
				bcvCell.setCellStyle(csPool.getCellStyle("0"));
				bcvCell.setCellFormula("SUMIF(AREA_BRIGHT_COLORED_VEGETABLE, 1, AREA_INTAKE)");

			}


		}

		return rowIndex;
	}










}
