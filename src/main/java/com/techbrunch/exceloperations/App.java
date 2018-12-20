package com.techbrunch.exceloperations;

import java.util.Scanner;

/**
 * ExcelOperationsApp
 * 
 * Developed by TechBrunch!
 *
 */
public class App {
	public static void main(String[] args) {
		System.out.println("Enter an option:\n");
		System.out.println("1) Write to Excel");
		System.out.println("2) Read from Excel");
		Scanner scannerObj = new Scanner(System.in);
		int userResponse = 0;
		try {
			userResponse = scannerObj.nextInt();
			scannerObj.close();
			if (!(userResponse == 1 || userResponse == 2)) {
				throw new RuntimeException("Provided input is not valid");
			}
		} catch (Exception ex) {
			System.out.println(ex.getMessage());
		}
		final AppUtil appUtil = new AppUtil();
		try {
			appUtil.performAction(userResponse);
		} catch (Exception ex) {
			ex.printStackTrace();
			System.out.println("An internal error occurred while processing the request!");
		}
	}
}
