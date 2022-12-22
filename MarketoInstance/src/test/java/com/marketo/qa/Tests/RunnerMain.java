package com.marketo.qa.Tests;


import java.util.ArrayList;
import java.util.List;

import org.testng.TestNG;





public class RunnerMain {
	
	static TestNG testng;
	public static void main(String[] args) throws Exception {
		
		testng =new TestNG();
		String path= System.getProperty("user.dir")+"//Config//testng.xml";
//		System.out.println(path);
			
		List<String> XmlList= new ArrayList<String>();
		XmlList.add(path);
		// TODO Auto-generated method stub
//		
//		testng.setTestClasses(new Class[] {FetchChampianCounts.class,FetchDesignStudioData.class,
//				FetchAdminDatas.class,FetchModelsCount.class,
//				FetchLeadsCount.class,FetchDataFromChangeScore.class});
		
		testng.setTestSuites(XmlList);	
		testng.run();
		
		
		

	}


}
