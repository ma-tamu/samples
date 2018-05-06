/**
 *
 */
package jp.co.project.venus.validation;

import java.lang.annotation.Annotation;

import jp.co.project.venus.enumeration.ValidationResultBean;
import jp.co.project.venus.model.BaseDto;

/**
 * @author M.Tamura
 *
 */
public abstract class ExcelRelationValidator<A extends Annotation, T extends BaseDto> {

	public void initialize(A annotation) {
	}

	public abstract ValidationResultBean validator(int rowIdx, T target);
}
