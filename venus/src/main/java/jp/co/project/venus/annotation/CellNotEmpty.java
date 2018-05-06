/**
 *
 */
package jp.co.project.venus.annotation;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import jp.co.project.venus.validation.CellNotEmptyValidator;

/**
 * セルの入力チェック
 *
 * @author M.Tamura
 *
 */
@Documented
@ExcelValidation(validatedBy = CellNotEmptyValidator.class)
@Retention(RUNTIME)
@Target(FIELD)
public @interface CellNotEmpty {

	String label();

	String message() default "未入力です。";
}
