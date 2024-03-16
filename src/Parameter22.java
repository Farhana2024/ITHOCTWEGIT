package Pakage1;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.Select;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

public class Parameter22 {

    public static void main(String[] args) throws IOException, InterruptedException {

        FileInputStream fis;
        FileOutputStream fos;
        XSSFWorkbook workbook;
        String filepath = "src/test/resources/dataforselenium.xlsx";

        fis = new FileInputStream(filepath);
        workbook = new XSSFWorkbook(fis);

        XSSFSheet sheet1 = workbook.getSheetAt(0);
       // sheet1.getRow(1).getCell(0).getStringCellValue();
        String url = sheet1.getRow(1).getCell(0).getStringCellValue();
        String sendto = sheet1.getRow(1).getCell(1).getStringCellValue();
        System.out.println(url);


        WebDriver driver = new EdgeDriver();
        driver.manage().window().maximize();
        driver.get(url);

        driver.findElement(By.linkText("Sample Forms")).click();


        // Implicit
        driver.manage().timeouts().implicitlyWait(50,TimeUnit.SECONDS);

        //Explicit


        if (sendto.equals("Marketing Department")) {
            driver.findElement(By.cssSelector("input[name='email_to[]'][value='0']")).click();

        } else if (sendto.equals("Sales")) {
            driver.findElement(By.cssSelector("input[name='email_to[]'][value='1']")).click();

        } else if (sendto.equals("Customer Service")) {
            driver.findElement(By.cssSelector("input[name='email_to[]'][value='2']")).click();

        } else {
            System.out.println("The given value is not available in the application");
        }




        String subject = sheet1.getRow(1).getCell(2).getStringCellValue();
        driver.findElement(By.id("subject")).sendKeys(subject);



        String email =sheet1.getRow(1).getCell(3).getStringCellValue();
        driver.findElement(By.id("email")).sendKeys(email);


        String testbox =sheet1.getRow(1).getCell(4).getStringCellValue();
        driver.findElement(By.id("q1")).sendKeys(testbox);




        String multibox = sheet1.getRow(1).getCell(5).getStringCellValue();
        driver.findElement(By.id("q2")).sendKeys(multibox);
        Thread.sleep(200);


        // dropdown box button need to work
        String dropdownval = sheet1.getRow(1).getCell(6).getStringCellValue();
        WebElement dropdown = driver.findElement(By.id("q3"));
        Select dropdowns1 = new Select(dropdown);
        dropdowns1.selectByValue(dropdownval);
        Thread.sleep(200);


        // Radio button

        String radiobt = sheet1.getRow(1).getCell(7).getStringCellValue();
        System.out.println(radiobt);
//        driver.findElement(By.xpath("//input[@name='q4'][@value='Third Option']")).click();
        driver.findElement(By.xpath("//input[@name='q4'][@value='" +radiobt.trim()+  "']")).click();
        Thread.sleep(200);


        // Check box multi answer

        String Ckboxmulti = sheet1.getRow(1).getCell(8).getStringCellValue();
        String checkboxmultiplearray[] = Ckboxmulti.split(",");
        for (int i=0; i< checkboxmultiplearray.length;i++) {
   //         driver.findElement(By.xpath("//input[@nme='checkbox6[]'][@value='"+checkboxmultiplearray[i].trim()+"']")).click();
            driver.findElement(By.xpath("//input[@name='checkbox6[]'][@value='"+checkboxmultiplearray[i].trim()+"']")).click();
        }

        Thread.sleep(200);


        // date Selector
        String dateselect = sheet1.getRow(1).getCell(9).getStringCellValue();
        driver.findElement(By.id("q7")).sendKeys(dateselect.trim());
        Thread.sleep(200);

        driver.findElement(By.id("q7")).sendKeys(Keys.ESCAPE);
        Thread.sleep(200);



        // Pre Defined US States
        String city = sheet1.getRow(1).getCell(10).getStringCellValue();
        WebElement state = driver.findElement(By.id("q8"));
        Select s2 = new Select(state);
        s2.selectByValue(city.trim());
        Thread.sleep(200);



        // Pre-Defined Field - Countries
        String cuntry = sheet1.getRow(1).getCell(11).getStringCellValue();
        WebElement name = driver.findElement(By.id("q9"));
        Select s3 = new Select(name);
        s3.selectByValue(cuntry.trim());
        Thread.sleep(200);



        // Canadean provence
        String canada = sheet1.getRow(1).getCell(12).getStringCellValue();
        WebElement provence = driver.findElement(By.id("q10"));
        Select s4 = new Select(provence);
        s4.selectByValue(canada.trim());
        Thread.sleep(200);


        // Pre Defined name
        String tittle = sheet1.getRow(1).getCell(13).getStringCellValue();
        WebElement tittle1 = driver.findElement(By.name("q11_title"));
        Select s5 = new Select(tittle1);
        s5.selectByValue(tittle.trim());
        Thread.sleep(200);

        // Pre Defined Firstname
        String Firstname = sheet1.getRow(1).getCell(14).getStringCellValue();
        driver.findElement(By.name("q11_first")).sendKeys(Firstname.trim());
        Thread.sleep(200);

        // Pre Defined Surname
        String Surname = sheet1.getRow(1).getCell(15).getStringCellValue();
        driver.findElement(By.name("q11_last")).sendKeys(Surname.trim());
        Thread.sleep(200);

        // Date of Birth - Month
        String Month = sheet1.getRow(1).getCell(16).getStringCellValue();
        WebElement Month1 = driver.findElement(By.name("q12_month"));
        Select s6 = new Select(Month1);
        s6.selectByValue(Month);
        Thread.sleep(300);
        // Screenshot1
        TakesScreenshot scrt3 = (TakesScreenshot) driver;
        File src3 = scrt3.getScreenshotAs(OutputType.FILE);
        FileUtils.copyFile(src3,new File("C:\\Users\\neham\\IdeaProjects\\ITHOCTSelenium2\\Screenshot\\DOBMONTH.jpg"));

        // Date of Birth - Day
        String Day = sheet1.getRow(1).getCell(17).getStringCellValue();
        WebElement day1 = driver.findElement(By.name("q12_day"));
        Select s7 = new Select(day1);
        s7.selectByValue(Day);
        Thread.sleep(300);
        // Screenshot2
        TakesScreenshot scrt2 = (TakesScreenshot) driver;
        File src2 = scrt2.getScreenshotAs(OutputType.FILE);
        FileUtils.copyFile(src2 , new File("C:\\Users\\neham\\IdeaProjects\\ITHOCTSelenium2\\Screenshot\\DOBday.jpg"));

        // Date of Birth - year

        String year = sheet1.getRow(1).getCell(18).getStringCellValue();
        WebElement year1 = driver.findElement(By.name("q12_year"));
        Select s8 = new Select(year1);
        s8.selectByValue(year);
        // Screenshot3
        TakesScreenshot scrt1 = (TakesScreenshot) driver;
        File src1 = scrt1.getScreenshotAs(OutputType.FILE);
        FileUtils.copyFile(src1, new File("C:\\Users\\neham\\IdeaProjects\\ITHOCTSelenium2\\Screenshot\\dateofbirth.jpg"));

        // File Attachment

        String Fileattach = sheet1.getRow(1).getCell(19).getStringCellValue();
        driver.findElement(By.name("attach4589")).sendKeys("C:\\Users\\neham\\Downloads\\"+Fileattach.trim()+".pdf" );
        TakesScreenshot skk = (TakesScreenshot) driver;
        File sk1= skk.getScreenshotAs(OutputType.FILE);
        FileUtils.copyFile(sk1,new File("C:\\Users\\neham\\IdeaProjects\\ITHOCTSelenium2\\Screenshot\\Attachfile.jpg"));


        // passing data from Selenium to excel
        Cell data1 = sheet1.getRow(1).createCell(21);
        data1.setCellValue("hello Farhana");

        Cell data2 = sheet1.getRow(1).createCell(22);
        String daata2 = driver.findElement(By.className("name")).getText();
        data2.setCellValue(daata2);

        Cell data3 = sheet1.getRow(1).createCell(23);
        String daata3 = driver.findElement(By.className("link")).getText();
        data3.setCellValue(daata3);

        Cell data4 = sheet1.getRow(1).createCell(24);
        String daata4 = driver.findElement(By.className("topbullet")).getText();
        data4.setCellValue(daata4);

        // code for excel file

        fos= new FileOutputStream(filepath);
        workbook.write(fos);
        workbook.close();
    }

}
