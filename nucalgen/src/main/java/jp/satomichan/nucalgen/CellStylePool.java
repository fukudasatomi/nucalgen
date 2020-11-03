package jp.satomichan.nucalgen;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;

public class CellStylePool {

	private Workbook workbook;

	CellStylePool(Workbook workbook){
		this.workbook = workbook;
	}


	private Map<String, CellStyle> cellStyleMap = new HashMap<String, CellStyle>();



	CellStyle getCellStyle(String format) {
		if(this.cellStyleMap.containsKey(format)) {
			return this.cellStyleMap.get(format);
		}else {
			CellStyle cs = this.workbook.createCellStyle();
			XSSFDataFormat xssfFormat = (XSSFDataFormat) this.workbook.createDataFormat();
			cs.setDataFormat(xssfFormat.getFormat(format));
			this.cellStyleMap.put(format, cs);

			return this.cellStyleMap.get(format);
		}
	}


}
