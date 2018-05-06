/**
 *
 */
package jp.co.project.venus.formula;

import java.util.List;

import jp.co.project.venus.annotation.SaveFlag;
import jp.co.project.venus.model.ExcelWorkbook;
import jp.co.project.venus.model.TransferReflectionBean;

/**
 * 更新フラグの数式作成クラス
 *
 * @author M.Tamura
 */
public class DefaultSaveFlagFormula extends CellFormula<SaveFlag> {

	private String formula;
	private String[] formulaFieldList;
	private static final String BEFORER_SHEET = "{beforeSheet}";

	@Override
	public void initialize(SaveFlag annotation) {
		this.formula = annotation.formula();
		this.formulaFieldList = annotation.formulaFieldList();
	}

	@Override
	public String formula(int rowIdx, List<TransferReflectionBean> reflectionFiledNameList, String... sheets) {
		int formulaIdx = 0;
		String workFormula = this.formula;

		for (String formulaField : formulaFieldList) {
			for (TransferReflectionBean reflectionBean : reflectionFiledNameList) {
				if (formulaField.equals(reflectionBean.getFiledName())) {
					workFormula = workFormula.replace("{" + formulaIdx + "}",
							ExcelWorkbook.toCellAddress(reflectionBean.getOffset(), rowIdx));
					formulaIdx++;
					break;
				}
			}
		}

		workFormula = workFormula.replace("{" + formulaFieldList.length + "}", sheets[0]);

		return workFormula;
	}

}
