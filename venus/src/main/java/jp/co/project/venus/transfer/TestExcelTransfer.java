/**
 *
 */
package jp.co.project.venus.transfer;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @author tamu
 *
 */
public class TestExcelTransfer extends ExcelService {

	/* (非 Javadoc)
	 * @see jp.co.project.venus.transfer.ExcelService#getSheetInfo()
	 */
	@Override
	protected Map<String, Integer> getSheetInfo() {
		Map<String, Integer> sheetMap = new LinkedHashMap<>();
		sheetMap.put("Sheet1", 4);
		return sheetMap;
	}

	@Override
	protected void makeFormula(int rowIdx) {
		// TODO 自動生成されたメソッド・スタブ

	}

	@Override
	protected void makeFormulaSaveFlg(int rowIdx) {
		// TODO 自動生成されたメソッド・スタブ

	}

	@Override
	protected int getNormalRow() {
		// TODO 自動生成されたメソッド・スタブ
		return 8;
	}

	@Override
	protected int getLockRow() {
		// TODO 自動生成されたメソッド・スタブ
		return 9;
	}

	@Override
	protected int getDetailRow() {
		// TODO 自動生成されたメソッド・スタブ
		return 12;
	}

	@Override
	protected String getInputSheetName() {
		// TODO 自動生成されたメソッド・スタブ
		return "入力シート";
	}

	@Override
	protected String getBeforeSheetName() {
		// TODO 自動生成されたメソッド・スタブ
		return "入力前シート";
	}

	@Override
	protected String getCodeTableSheetName() {
		// TODO 自動生成されたメソッド・スタブ
		return null;
	}


}
