package util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Gurugram_gov {

	public static void main(String[] args) throws IOException, InterruptedException {
		
		File file = new File("/Users/mownicabodala/Documents/Selenium Workspace/MissionHumane/src/main/resources/datafiles/ExtractedWebTable79.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sh = wb.createSheet("covid");
		System.setProperty("webdriver.chrome.driver","/Users/mownicabodala/Desktop/Drivers/chromedriver");
		WebDriver driver= new ChromeDriver();
		
		driver.get("http://covidggn.com/public/pages/gurugram-hospitals");
		driver.manage().window().maximize();
		Thread.sleep(3000);
		
		 sh.createRow(0).createCell(1).setCellValue("Hospital Name");
	     sh.getRow(0).createCell(2).setCellValue("Vacant");  
	     sh.getRow(0).createCell(3).setCellValue("Contact");   
	     sh.getRow(0).createCell(4).setCellValue("Last updated on");   
	    
		
		WebElement table = driver.findElement(By.xpath("//table[@class='table hospitalTable table-bordered']"));

		List<WebElement> totalRows = table.findElements(By.tagName("tr"));
		for(int row=1; row<totalRows.size(); row=row+2)
		{
			XSSFRow rowValue = sh.createRow(row);
			List<WebElement> totalColumns = totalRows.get(row).findElements(By.tagName("td"));
			for(int col=1; col<totalColumns.size(); col++)
			{
				String cellValue = totalColumns.get(col).getText();
				System.out.print(cellValue + "\t");
				rowValue.createCell(col).setCellValue(cellValue);
			}
			System.out.println();
		}
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
		wb.close();
	}
	
		
		
	}


