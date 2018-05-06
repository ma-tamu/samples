/**
 *
 */
package jp.co.project.venus.validation;

import org.apache.commons.lang3.StringUtils;

import jp.co.project.venus.annotation.CellNotEmpty;
import jp.co.project.venus.enumeration.ValidationResultBean;
import jp.co.project.venus.model.ExcelWorkbook;

/**
 * @author M.Tamura
 *
 */
public class CellNotEmptyValidator extends ExcelValidator<CellNotEmpty, String> {

	private String label;
	private String message;

	@Override
	public void initialize(CellNotEmpty annotation) {
		this.label = annotation.label();
		this.message = annotation.message();
	}

	@Override
	public ValidationResultBean validator(int colIdx, int rowIdx, String value) {
		ValidationResultBean resultBean = null;
		if (StringUtils.isEmpty(value)) {
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
