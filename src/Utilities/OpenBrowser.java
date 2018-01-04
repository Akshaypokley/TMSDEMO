package Utilities;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

/**
 * Created by AKSHAY on 26/04/2017.
 */
public class OpenBrowser {

    static WebDriver driver;
    public static WebDriver openBrowser(String browserNm)
    {
        switch (browserNm)
        {
            case "chrome":
                System.setProperty("webdriver.chrome.driver","Driver/chromedriver.exe");
                driver=new ChromeDriver();
                driver.manage().window().maximize();
                break;

            case "Firefox":
                System.setProperty("webdriver.gecko.driver","Driver/geckodriver.exe");
                driver=new FirefoxDriver();
                driver.manage().window().maximize();
                break;

            case "IE":
                System.setProperty("webdriver.IE.driver","Driver/IEDriverServer.exe");
                driver=new FirefoxDriver();
                driver.manage().window().maximize();
                break;

                default:
                    System.out.println("browser : " + browserNm + " is invalid, Launching Firefox as browser of choice..");
                    System.setProperty("webdriver.chrome.driver","Driver/chromedriver.exe");
                    driver=new ChromeDriver();
                    driver.manage().window().maximize();
                    break;


        }

        return driver;
    }


    public static void  GetUrl(String URL)

    {

        driver.get("https://www.rrcw.info/pp/adminlte/full/admindashboard.html");
}
}
