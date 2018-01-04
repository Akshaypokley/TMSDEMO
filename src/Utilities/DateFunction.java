package Utilities;

import org.joda.time.DateTime;
import org.joda.time.Months;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;


/**
 * Created by AKSHAY on 20/04/2017.
 */
public class DateFunction {

 static WebDriver driver;
    public static void DateFun(WebDriver driver, String seDate) throws ParseException

    {
        //driver=openbrowser("chrome");
        //driver.findElement(By.xpath("html/body/div[1]/aside/section/ul/li[5]/a/span")).click();


        SimpleDateFormat myDateFormat = new SimpleDateFormat("dd/MM/YYYY");

        SimpleDateFormat calDateFormat = new SimpleDateFormat("MMMM yyyy");

        Date setDate=myDateFormat.parse(seDate);

        driver.findElement(By.xpath(".//*[@id='radPossessionDate_popupButton']")).click();
        driver.manage().timeouts().implicitlyWait(40,TimeUnit.SECONDS);

        Date curDate = calDateFormat.parse(driver.findElement(By.xpath("//html/body/div/div[4]/table/thead/tr/td/table/tbody/tr/td[3]")).getText());
        System.out.println(curDate);

        int monthDiff = Months.monthsBetween(new DateTime(curDate).withDayOfMonth(1),new DateTime(setDate).withDayOfMonth(1)).getMonths();
        boolean isFuture = true;
        System.out.println(monthDiff);
        // decided whether set date is in past or future
        if(monthDiff <0){
            isFuture = false;
            monthDiff*=-1;
        }
        // iterate through month difference
        for(int i=1;i<=monthDiff;i++){
            driver.findElement(By.xpath(".//*[@id='radPossessionDate_calendar']/thead/tr/td/table/tbody/tr/td/a[@class="+ (isFuture ? "'rcNext'" : "'rcPrev'") + "]")).click();
        }
        // Click on Day of Month from table
       /* for(int j=1;j<7;j++){
*/
            driver.findElement(By.xpath(".//*[@id='radPossessionDate_calendar_Top']/tbody/tr[2]/td/a[text()='" + (new DateTime(setDate).getDayOfMonth()) + "']")).click();

     /*   }*/


    }


}
