/**
 *
 */
package jp.co.project.venus.validation;

import java.lang.annotation.Annotation;

import jp.co.project.venus.enumeration.ValidationResultBean;

/**
 * @author M.Tamura
 *
 */
public abstract class ExcelValidator<A extends Annotation, T> {

	/**
	 * 初期処理
	 *
	 * @param annotation
	 *            Annotation
	 */
	public void initialize(A annotation) {
	};

	public abstract ValidationResultBean validator(int colIdx, int rowIdx, T target);
}
