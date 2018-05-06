/**
 *
 */
package jp.co.project.venus.validation;

import org.apache.commons.lang3.StringUtils;

import jp.co.project.venus.annotation.CellNumber;
import jp.co.project.venus.enumeration.ValidationResultBean;
import jp.co.project.venus.model.ExcelWorkbook;

/**
 * @author M.Tamura
 *
 */
public class CellNumberValidator extends ExcelValidator<CellNumber, String> {

	private String label;
	private String message;
	private static final String NUMBER_REG = "^-?[0-9]*.?[0-9]+$";

	@Override
	public void initialize(CellNumber annotation) {
		this.label = annotation.label();
		this.message = annotation.message();
	}

	@Override
	public ValidationResultBean validator(int colIdx, int rowIdx, String value) {
		ValidationResultBean resultBean = null;
		if (StringUtils.isEmpty(value) || !value.matches(NUMBER_REG)) {
			resultBean = new ValidationResultBean();
			resultBean.setColIdx(colIdx);
			resultBean.setRowIdx(rowIdx);
			resultBean.setCellAddress(ExcelWorkbook.toCellAddress(colIdx, rowIdx));
			resultBean.setLabel(this.label);
			resultBean.setMessage(this.message);
		}
		return resultBean;
	}

}
