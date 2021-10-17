/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

/**
 * General exception class for EEH.
 * @author jhorvath
 */
public class EEHException extends Exception {

	private static final long serialVersionUID = 5636424200555163709L;

	public EEHException(String string) {
		super(string);
	}

	public EEHException(String string, Exception ex) {
		super(string, ex);
	}

}
