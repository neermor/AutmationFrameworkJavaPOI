package com.marketo.qa.Tests;

import com.marketo.qa.base.TestBase;

public class test extends TestBase {

	public static void main(String Args[]) throws Exception {
		/// String ExcelPath=System.getProperty("user.dir")+"\\testng.xml";
		// System.out.println(ExcelPath);
		new TestBase().OpenSupportTool();
	}

}
