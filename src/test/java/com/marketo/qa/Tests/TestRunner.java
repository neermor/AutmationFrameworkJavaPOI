package com.marketo.qa.Tests;

import java.util.ArrayList;
import java.util.List;

import org.testng.TestNG;

public class TestRunner {

	static TestNG testng;
	public static void main(String[] args) {
		testng =new TestNG();
		String path= System.getProperty("user.dir")+"testng.xml";
		List<String> XmlList= new ArrayList<String>();
		XmlList.add(path);
		// TODO Auto-generated method stub
		
		testng.setTestClasses(new Class[] {FetchDataFromChangeScore.class});
		
		testng.run();
		

	}

}
