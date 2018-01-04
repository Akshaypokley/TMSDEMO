
import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.read.biff.BiffException;
import jxl.write.*;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.joda.time.Seconds;
import org.junit.rules.Timeout;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.*;

import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;

import static Utilities.AttachFunction.AttachFuntn;
import static Utilities.OpenBrowser.GetUrl;
import static Utilities.OpenBrowser.openBrowser;
import static jxl.format.Colour.*;
import static jxl.format.Colour.LIGHT_TURQUOISE;

/**
 * Created by AKSHAY on 04/01/2018.
 */
public class tmsDemo {

    static WebDriver driver;
    public Label l4;
    public static WritableCellFormat cellFormat;
    public static WritableCellFormat cellFormat1;
    public static WritableCellFormat cellFormat3;
    public static WritableCellFormat cellFormat4;
    public WritableCellFormat cellFormat2;
    public static WritableCellFormat cellFormat6;
    public static WritableCellFormat cellFormat5;
    public WritableWorkbook writableTempSource;
    public WritableWorkbook copyDocument;
    public WritableSheet sourceSheet;
    public static WritableSheet targetSheet;
    public Workbook sourceDocument;
    /*****************************************************************/
    private static int n = 2;
    private static int j = 1;
    public static String Result;
    public static String Actual;
    static int LastRow;
    static int SetBord;
    static final java.util.regex.Pattern String = java.util.regex.Pattern.compile("^[A-Za-z, ]++$");
    public static int k1=1;

    @BeforeTest
    public  void OutputExcelCreation() throws IOException, BiffException, WriteException {

        sourceDocument = Workbook.getWorkbook(new File("ExcelData/InputData/TMSData.xls"));
        writableTempSource = Workbook.createWorkbook(new File("ExcelData/InputData/temp.xls"), sourceDocument);
        copyDocument = Workbook.createWorkbook(new File("ExcelData/Result/TMSTestReport.xls"));
        sourceSheet = writableTempSource.getSheet(0);
        targetSheet = copyDocument.createSheet("sheet 1", 2);

        WritableFont cellFont = new WritableFont(WritableFont.COURIER, 11);
        cellFont.setBoldStyle(WritableFont.BOLD);
/************************************************************************************************/
        WritableFont cellFont2 = new WritableFont(WritableFont.COURIER, 10);
        cellFont2.setColour(BLACK);
        //cellFont2.setBoldStyle(WritableFont.BOLD);
        cellFormat1 = new WritableCellFormat(cellFont2);
        cellFormat1.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        cellFormat1.setWrap(true);
/*******************************************************************************************************/
/************************************************************************************************/
        WritableFont cellFont3 = new WritableFont(WritableFont.COURIER, 10);
        cellFont3.setColour(RED);
        // cellFont3.setBoldStyle(WritableFont.BOLD);
        cellFormat3 = new WritableCellFormat(cellFont3);
        cellFormat3.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        cellFormat3.setWrap(true);

        WritableFont cellFont4 = new WritableFont(WritableFont.COURIER, 10);
        cellFont4.setColour(GREEN);
        // cellFont4.setBoldStyle(WritableFont.BOLD);
        cellFormat4 = new WritableCellFormat(cellFont4);
        cellFormat4.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        cellFormat4.setWrap(true);


        cellFormat = new WritableCellFormat(cellFont);
        cellFormat.setBackground(LIGHT_BLUE);
        cellFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        cellFormat.setWrap(true);
        cellFormat2 = new WritableCellFormat(cellFont);
        cellFormat2.setBackground(RED);
        //cellFormat.setAlignment(jxl.format.Alignment.getAlignment(20));
        WritableFont cellFont5 = new WritableFont(WritableFont.COURIER, 18);
        cellFont5.setColour(BLACK);
        cellFont5.setBoldStyle(WritableFont.BOLD);
        cellFormat5 = new WritableCellFormat(cellFont5);
        cellFormat5.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        cellFormat5.setBackground(LIGHT_BLUE);
        cellFormat5.setAlignment(Alignment.CENTRE);


        cellFormat6 = new WritableCellFormat(cellFont2);
        cellFormat6.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        cellFormat6.setWrap(true);
        cellFormat6.setBackground(LIGHT_TURQUOISE);
        //  sheet.addCell(new Label(col, 1, "CCCCC", cellFormat));

        for (int row = 0; row < sourceSheet.getRows(); row++) {
            for (int col = 0; col < sourceSheet.getColumns(); col++) {
                WritableCell readCell = sourceSheet.getWritableCell(col, row);
                WritableCell newCell = readCell.copyTo(col, row);
                CellFormat readFormat = readCell.getCellFormat();

                WritableCellFormat newFormat = new WritableCellFormat(readFormat);
                newCell.setCellFormat(newFormat);
                targetSheet.addCell(newCell);


                Label l2=new Label(5,1,"Actual",cellFormat);

                Label l3=new Label(6,1,"Status",cellFormat);
                //Label l4=new Label(4,row,"",cellFormat);
                int widthInChars = 36;
                int widthInChars2 = 18;
                int widthInChars1 = 16;
                targetSheet.setColumnView(2, widthInChars1);
                targetSheet.setColumnView(3, widthInChars1);
                targetSheet.setColumnView(1, widthInChars1);
                targetSheet.setColumnView(4, widthInChars);
                targetSheet.setColumnView(5, widthInChars);

/*-----------------------------------------------------------------------------------------------------------------------*/
                targetSheet.setColumnView(0, widthInChars2);
                targetSheet.setColumnView(2, widthInChars2);

                targetSheet.setColumnView(3, widthInChars2);
                targetSheet.mergeCells(0, 0, 6, 0);
                Label lable = new Label (0, 0,
                        "Add Assignment screen test  report",cellFormat5);
                targetSheet.addCell(lable);
                targetSheet.addCell(l2);
                targetSheet.addCell(l3);
                //targetSheet.addCell(l4);

            }
        }

    }
    @AfterTest
    public void f() throws IOException, WriteException
    {

        copyDocument.write();
        copyDocument.close();
        writableTempSource.close();
        sourceDocument.close();

    }


    @Test(dataProvider = "hybridData")
    public static void TMSTest(String testcaseName, String keyword, String objectName, String value, String Expected) throws Exception {

        if (testcaseName != null && testcaseName.length() != 0) {

            driver = openBrowser("chrome");
            GetUrl("url");
            Thread.sleep(200);
            SetBord = j++;
            Label l7 = new Label(5, SetBord, "", cellFormat6);
            targetSheet.addCell(l7);
            Label l8 = new Label(6, SetBord, "", cellFormat6);
            targetSheet.addCell(l8);
        } else {
            SetBord = j++;
        }

        try {
            switch (keyword.toUpperCase()) {

                case "SUBMISSION":
                    switch(objectName)
                    {
                        case "Add Assignment":
                            driver.findElement(By.xpath(".//*[@id='submit']")).click();
                            Result="pass";
                            break;

                    }

                case "ASSIGN":
                    switch (objectName)
                    {
                        case "Online Managers":
                            driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[1]/div[1]/div/div[1]/label/input")).click();
                            WebElement Bank=driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[1]/div[1]/div/div[2]/select"));
                            Select combo1=new Select(Bank);
                            combo1.selectByVisibleText(value);

                            break;

                        case "Offline Managers":
                            driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[1]/div[2]/div/div[1]/label/input")).click();
                            WebElement om=driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[1]/div[2]/div/div[2]/select"));
                            Select combo2=new Select(om);
                            combo2.selectByVisibleText(value);

                            break;
                        case "Online Writers":
                            driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[2]/div[1]/div/div[1]/label/input")).click();
                            WebElement ow=driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[1]/div[2]/div/div[2]/select"));
                            Select combo3=new Select(ow);
                            combo3.selectByVisibleText(value);
                            break;
                        case "Offline Writers":
                            driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[2]/div[2]/div/div[1]/label/input")).click();
                            WebElement ow1=driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[2]/div[2]/div/div[2]/select"));
                            Select combo=new Select(ow1);
                            combo.selectByVisibleText(value);
                            break;

                        case "Online ProofReaders":

                            driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[3]/div[1]/div/div[1]/label/input")).click();
                            WebElement ow11=driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[3]/div[1]/div/div[2]/select"));
                            Select combo11=new Select(ow11);
                            combo11.selectByVisibleText(value);
                            break;
                        case "Offline ProofReaders":
                            driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[3]/div[2]/div/div[1]/label/input")).click();
                            WebElement ow111=driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[3]/div[2]/div/div[2]/select"));
                            Select combo111=new Select(ow111);
                            combo111.selectByVisibleText(value);
                            break;

                    }
                        break;
                case "ATTACH":
                    switch (objectName)
                    {
                        case "Client Attachments":
                            driver.findElement(By.xpath(".//*[@id='file']")).click();
                            Thread.sleep(600);
                            AttachFuntn(driver, "G:\\The.docx");
                            Result="pass";

                    }
                            break;

                case "MENU":

                    switch (objectName) {



                        case "Click on menu":
                            try {
                                List<WebElement> cells = driver.findElements(By.xpath("/html/body/div[1]/aside[1]/section/ul[2]/li/a/span"));

                                for (WebElement cell : cells) {
                                    String fiels = cell.getText();
                                    System.out.println(fiels);
                                    if (fiels.equals(value)) {
                                        System.out.println(fiels);

                                            cell.click();
                                            System.out.println("Value Name is :-***" + value + "***");
                                            Result = "pass";
                                            break;

                                    } else Result="fail";
                                        Actual="Tab not present.";
                                    }   ++k1;


                            } catch (Throwable e) {

                            }
                            break;

                    }    break;

                case "SETTEXT":

                    switch (objectName) {
                        case "Enter Reason":
                            driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[4]/div/div[2]/div/div[4]/div/div/textarea")).sendKeys(value);
                            Result="pass";break;

                        case "Number of Words":
                        driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[1]/div[7]/div/input")).sendKeys(value);
                            Result="pass";break;

                        case "Description":
                            driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[1]/div[9]/div/textarea")).sendKeys(value);
                            Result="pass"; break;


                        case "Name":
                            WebElement name=driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[1]/div[2]/div/input"));
                            name.clear();
                            name.sendKeys(value);
                            final String NM = name.getAttribute("value");


                            System.out.println(NM);
                            if(NM.isEmpty())
                            {
                                try {

                                    driver.manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);
                                    WebElement validationtext= driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[1]/div[2]/div/span"));
                                    String vatext=validationtext.getText();
                                    if(vatext.equals(Expected)) {
                                        Result = "pass";
                                    }else {

                                        Result="fail";
                                        Actual=vatext;

                                    }
                                }catch (NoSuchElementException j)
                                {
                                    Result="fail";
                                    Actual=j.getMessage();
                                }

                            }
                            else{

                                if (NM.equals(value)) {
                                    if (!String.matcher(NM).matches()) {
                                        try {

                                                if (Actual.equals(Expected)) {
                                                    Result = "pass";
                                                } else {
                                                    Result = "Fail";
                                                }
                                                System.out.println(Actual);
                                                //    Thread.sleep(50);


                                        } catch (Throwable e) {
                                            Actual = "Alert message not display .";
                                            Result = "Fail";
                                        }
                                    } else {
                                        Result = "pass";
                                        System.out.println(NM);
                                        System.out.println(Result);
                                    }
                                } else {
                                    if (Actual.equals(Expected)) {
                                        Result = "pass";
                                    } else {
                                        Result = "Fail";
                                    }
                                }




                    }


           }break;

                case "SELECT":

                    switch (objectName) {

                        case "Deadline":
                            WebElement dateBox = driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[1]/div[4]/div/div[1]/div/input"));
                            dateBox.sendKeys("0922018");
                            Result="pass";

                        break;

                        case "Client ID":
                            try {
                                Thread.sleep(2000);
                                WebElement Bank=driver.findElement(By.xpath("./*//*[@id='add-assign']/div/div[1]/div[3]/div/select"));
                                Select combo1=new Select(Bank);
                                combo1.selectByVisibleText(value);
                                Result="pass";
                            }catch (Throwable j)
                            {
                                System.out.println(j.getMessage());
                            }

                            break;

                        case "Time":
                            Thread.sleep(2000);
                            WebElement d2=driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[1]/div[4]/div/div[2]/div/input"));

                            d2.sendKeys("0111PM");Result="pass";
                            break;

                        case "Type":
                            try {
                                Thread.sleep(2000);
                                WebElement Bank=driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[1]/div[5]/div/select"));
                                Select combo1=new Select(Bank);
                                combo1.selectByVisibleText(value);  Result="pass";
                            }catch (Throwable j)
                            {
                                System.out.println(j.getMessage());
                            }
                            break;


                        case "Niche":
                            try {
                                Thread.sleep(2000);
                                WebElement Bank=driver.findElement(By.xpath(".//*[@id='add-assign']/div/div[1]/div[8]/div/select"));
                                Select combo1=new Select(Bank);
                                combo1.selectByVisibleText(value);  Result="pass";
                            }catch (Throwable j)
                            {
                                System.out.println(j.getMessage());
                            }
                            break;




                        }break;




                                default:

                    break;
            }


            if (testcaseName.isEmpty()) {
                LastRow = n++;
                if (Result.equals("pass")) {
                    Label l5 = new Label(5, LastRow, "As Exptected", cellFormat1);
                    targetSheet.addCell(l5);
                    Label l6 = new Label(6, LastRow, "PASS", cellFormat4);
                    targetSheet.addCell(l6);
                } else {

                    Label l5 = new Label(5, LastRow, Actual, cellFormat1);
                    targetSheet.addCell(l5);
                    Label l6 = new Label(6, LastRow, "FAIL", cellFormat3);
                    targetSheet.addCell(l6);
                }
            } else {
                LastRow = n++;

            }
        } catch (NullPointerException e) {
        }


    }

    @DataProvider(name = "hybridData")
    public Object[][] getDataFromDataprovider() throws IOException {
        Object[][] object = null;
        FileInputStream fis = new FileInputStream("ExcelData/InputData/TMSData.xls");
        HSSFWorkbook wb = new HSSFWorkbook(fis);
        HSSFSheet sh = wb.getSheet("Applicant");
        //  HSSFRow rows = sh.getRow(1);
//Read keyword sheet
//Find number of rows in Expl.excel file
        int rowCount = sh.getLastRowNum() - sh.getFirstRowNum();
        System.out.println(rowCount);
        object = new Object[rowCount][5];
        for (int i = 1; i < rowCount; i++) {

            HSSFRow row = sh.getRow(i + 1);


            for (int j = 0; j < row.getLastCellNum(); j++) {
//                System.out.println(row.getCell(j).toString());
                object[i][j] = row.getCell(j).toString();

            }


        }
        return object;


    }


}