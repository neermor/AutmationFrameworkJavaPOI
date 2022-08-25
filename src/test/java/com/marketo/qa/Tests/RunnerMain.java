package com.marketo.qa.Tests;

import java.util.ArrayList;
import java.util.List;

import org.testng.TestNG;





public class RunnerMain {
	
	static TestNG testng;
	public static void main(String[] args) throws Exception {
		
		testng =new TestNG();
		String path= System.getProperty("user.home")+"/Desktop/Config/testng.xml";
		System.out.println(path);
		List<String> XmlList= new ArrayList<String>();
		XmlList.add(path);
		// TODO Auto-generated method stub
		
//		testng.setTestClasses(new Class[] {FetchDataFromChangeScore.class,FetchAdminDatas.class,
//				FetchDesignStudioData.class,FetchLeadsCount.class,
//				FetchModelsCount.class});
		testng.setTestSuites(XmlList);
		testng.run();
		
		
		

	}


}
