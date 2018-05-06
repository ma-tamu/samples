/**
 *
 */
package jp.co.project.venus.formula;

import java.lang.annotation.Annotation;
import java.util.List;

import jp.co.project.venus.model.TransferReflectionBean;

/**
 * @author tamura
 *
 */
public abstract class CellFormula<A extends Annotation> {

	/**
	 * 初期処理
	 *
	 * @param annotation
	 *            Annotation
	 */
	public void initialize(A annotation) {
	};

	public abstract String formula(int rowIdx, List<TransferReflectionBean> reflectionFiledNameList, String... sheets);
}
