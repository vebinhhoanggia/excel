/**
 *
 */
package com.alxvn.backlog.handle;

/**
 * @author KEDD
 *
 */
public class IncorrectFullNameException extends Exception {
	/**
	 *
	 */
	private static final long serialVersionUID = -1163568384357303856L;

	public IncorrectFullNameException(String errorMessage) {
		super(errorMessage);
	}

	public IncorrectFullNameException(String errorMessage, Throwable err) {
		super(errorMessage, err);
	}
}
