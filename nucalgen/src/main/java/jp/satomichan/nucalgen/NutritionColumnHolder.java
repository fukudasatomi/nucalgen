package jp.satomichan.nucalgen;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.configuration.XMLConfiguration;

public class NutritionColumnHolder {

	private List<NutritionColumn> nutritionColumnList;

	List<NutritionColumn> getNutritionColumnList() {
		return this.nutritionColumnList;
	}

	void addNutritionColumn(NutritionColumn aNutritionColumn) {
		this.nutritionColumnList.add(aNutritionColumn);
	}

	NutritionColumnHolder(XMLConfiguration aConfig){
		this.nutritionColumnList = new ArrayList<NutritionColumn>();

		List<Object> names = aConfig.getList("cols.column.name");
		List<Object> dispNames = aConfig.getList("cols.column.disp_name");
		List<Object> aliases = aConfig.getList("cols.column.alias");
		List<Object> formats = aConfig.getList("cols.column.format");
		List<Object> units = aConfig.getList("cols.column.unit");
		List<Object> useRawValue = aConfig.getList("cols.column.use_raw_value");
		List<Object> useSum = aConfig.getList("cols.column.use_sum");

		for (Object aName : names) {
			NutritionColumn nc = new NutritionColumn();
			nc.setName((String) aName);
			nc.setDispName((String) dispNames.get(names.indexOf(aName)));
			nc.setAlias((String) aliases.get(names.indexOf(aName)));
			nc.setFormat((String) formats.get(names.indexOf(aName)));
			nc.setUnit((String) units.get(names.indexOf(aName)));
			nc.setUseRawValue(((String)useRawValue.get(names.indexOf(aName))).equalsIgnoreCase("true"));
			nc.setUseSum(((String)useSum.get(names.indexOf(aName))).equalsIgnoreCase("true"));

			this.addNutritionColumn(nc);
		}

	}


	public String toString() {
		String ret = "";
		for(NutritionColumn aColumn : this.getNutritionColumnList()) {
			ret += aColumn + "\n";
		}

		return ret;
	}



}
