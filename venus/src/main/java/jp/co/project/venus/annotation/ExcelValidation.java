/**
 *
 */
package jp.co.project.venus.annotation;

import static java.lang.annotation.ElementType.ANNOTATION_TYPE;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import jp.co.project.venus.validation.ExcelValidator;

/**
 * @author M.Tamura
 *
 */
@Documented
@Target({ ANNOTATION_TYPE })
@Retention(RUNTIME)
public @interface ExcelValidation {

	/**
	 * バリデートクラス
	 *
	 * @return Class<? extends ExcelValidator<?, ?>>[]
	 */
	Class<? extends ExcelValidator<?, ?>>[] validatedBy();
}
