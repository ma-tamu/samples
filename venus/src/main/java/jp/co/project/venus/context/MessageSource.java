/**
 *
 */
package jp.co.project.venus.context;

import java.util.Locale;
import java.util.ResourceBundle;

/**
 * @author M.Tamura
 *
 */
public class MessageSource {

	private static MessageSource instance = null;
	private ResourceBundle bundle = null;

	/**
	 * コンストラクタ
	 */
	private MessageSource() {
		this(Locale.JAPAN);
	}

	/**
	 * コンストラクタ
	 *
	 * @param locale
	 *            Locale
	 */
	private MessageSource(Locale locale) {
		bundle = ResourceBundle.getBundle("ValidatorMessages", locale);
	}

	/**
	 * メッセージリソースインスタンスを取得
	 *
	 * @return MessageSource
	 */
	public static MessageSource getInstance() {

		if (instance == null) {
			instance = new MessageSource();
		}

		return instance;
	}

	/**
	 * メッセージリソースインスタンスを取得
	 *
	 * @param locale
	 *            Locale
	 * @return MessageSource
	 */
	public static MessageSource getInstance(Locale locale) {

		if (instance == null) {
			instance = new MessageSource(locale);
		}

		return instance;
	}

	/**
	 *
	 * @param key String
	 * @param objects MessageSource
	 * @return String
	 */
	public String getMessage(String key, Object... objects) {
		int argIdx = 0;
		String message = null;
		final String argFormat = "{%d}";

		// プロパティからメッセージを取得
		message = bundle.getString(key);

		for (Object arg : objects) {
			message.replace(String.format(argFormat, argIdx), arg.toString());
			argIdx++;
		}

		return message;
	}
}
