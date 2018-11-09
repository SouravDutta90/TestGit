package DataDriven_LC;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;

import org.openqa.selenium.WebDriver;

import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;

import org.testng.annotations.BeforeMethod;

import org.testng.annotations.Test;

import org.testng.annotations.DataProvider;

import DataDriven_LC.ExcelUtils;

public class WebTest {
	WebDriver driver;

    
	@Test(dataProvider="shipmentData")
	
	public void testShipment(String id,String Name,String Weight,String ArrivalPort, String DepartuerPort,String Profit)throws  Exception{
	   
		String a=String.format("%-15s%-25s%-25s%-25s%-15s%15s",id,Name,Weight,ArrivalPort,DepartuerPort,Profit);
		System.out.print(a+"\n");  
		   
	  
	
		}
	
	@DataProvider
	
	public Object[][] shipmentData() throws Exception{
	
	     Object[][] testObjArray = ExcelUtils.getTableArray("D:\\SFDC.NEXT\\New Eclipse\\LC_DataDriven_2\\Shipments.xlsx","Sheet1");
	
	     return (testObjArray);
	
		}
	
	
}
