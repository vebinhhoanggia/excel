/**
 *
 */
package com.alxvn.backlog.util;

import com.alxvn.backlog.dto.BacklogDetail;
import com.alxvn.backlog.handle.IncorrectFullNameException;

/**
 *
 */
public class BacklogHelper {

	private BacklogHelper() {
		throw new IllegalStateException("Utility class");
	}

	public static BacklogDetail parseCsvRecord(String[] row) throws IncorrectFullNameException {
		// Accessing values by Header names

		return new BacklogDetail.Builder()
				/**/
				.build();
	}
}
