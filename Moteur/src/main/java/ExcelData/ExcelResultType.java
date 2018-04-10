package ExcelData;

public enum ExcelResultType {
	RESERVE_BEGIN_OF_PERIOD("RESERVE_BEGIN_OF_PERIOD"),
	RESERVE_END_OF_PERIOD("RESERVE_END_OF_PERIOD"),
	RISK_PREMIUM("RISK_PREMIUM"),
	PAID_PREMIUM("PAID_PREMIUM"),
	ENSURED_CAPITAL("ENSURED_CAPITAL"),
	RESERVE_PROFITSHARING("RESERVE_PROFITSHARING");
	

	private final String desc;
	
	private ExcelResultType(String description) {
		desc = description;
	}

	public String getDesc() {
		return desc;
	}
	
	
	
}
