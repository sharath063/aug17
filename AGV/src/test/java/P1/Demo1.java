package P1;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Demo1 {

	@Test
	public void test()
	{
		
	//	WebDriverManager.chromedriver().setup();
	//	WebDriver driver=new ChromeDriver();
		WebDriverManager.firefoxdriver().setup();
		WebDriver driver=new FirefoxDriver();
		driver.get("https://www.google.co.in");
		driver.close();
		WebDriverManager.firefoxdriver().setup();
	}

}
