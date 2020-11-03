package jp.satomichan.nucalgen;

public class NutritionColumn {
	private String name;
	private String dispName;
	private String format;
	private String unit;
	private boolean useRawValue;


	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public void setName(Object name) {
		this.name = (String) name;
	}


	public String getDispName() {
		return dispName;
	}

	public void setDispName(String disp_name) {
		this.dispName = disp_name;
	}


	public String getFormat() {
		return format;
	}

	public void setFormat(String format) {
		this.format = format;
	}


	public String getUnit() {
		return unit;
	}

	public void setUnit(String unit) {
		this.unit = unit;
	}

	public boolean isUseRawValue() {
		return useRawValue;
	}

	public void setUseRawValue(boolean useRawValue) {
		this.useRawValue = useRawValue;
	}


	public String toString() {
		String ret = "name={" + name + "} disp_name={" + dispName +
				"} format={" + format + "} unit={" + unit + "} useRawValue={" + useRawValue + "}";

		return ret;
	}


}
