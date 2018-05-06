/**
 *
 */
package jp.co.project.venus.annotation;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import jp.co.project.venus.validation.CellNumberValidator;


@Documented
@ExcelValidation(validatedBy = CellNumberValidator.class)
@Retention(RUNTIME)
@Target(FIELD)
/**
 * 数値チェック
 *
 * @author M.Tamura
 *
 */
public @interface CellNumber {

	String label();

	String message() default "数値を入力してください。";
}
