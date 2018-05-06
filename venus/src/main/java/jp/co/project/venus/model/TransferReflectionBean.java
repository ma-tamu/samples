/**
 *
 */
package jp.co.project.venus.model;

import lombok.Getter;
import lombok.Setter;

/**
 * @author M.Tamura
 *
 */
public class TransferReflectionBean {

	/**
	 * リフレクションフィールド名
	 */
	@Getter
	@Setter
	private String filedName;

	/**
	 * リフレクションフィールド名のオフセット
	 */
	@Getter
	@Setter
	private int offset;

	/**
	 * リフレクション時に設定する値
	 */
	@Getter
	@Setter
	private String filedValue;
}
