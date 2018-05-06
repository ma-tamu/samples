/**
 *
 */
package jp.co.project.venus.annotation;

import static java.lang.annotation.ElementType.ANNOTATION_TYPE;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import jp.co.project.venus.formula.CellFormula;

/**
 * @author M.Tamura
 *
 */
@Documented
@Target({ ANNOTATION_TYPE })
@Retention(RUNTIME)
public @interface ExcelFormula {

	/**
	 * バリデートクラス
	 *
	 * @return Class<? extends ExcelFormula<?, ?>>
	 */
	Class<? extends CellFormula<?>> formulaBy();

}
