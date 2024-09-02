import static org.junit.jupiter.api.Assertions.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.text.SimpleDateFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

class CreateMajorTest {

    @BeforeAll
    static void setUpBeforeClass() throws Exception {
        // Set up system property for ChromeDriver
        System.setProperty("webdriver.chrome.driver", "D:/chromedriver.exe");
    }

    @Test
    void testCreateMajorTest() throws Exception {
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
        Date date = new Date(0);
        String testDate = formatter.format(date);
        String testerName = "Titadach Sratongaun";

        String path = "C:/Users/HP-NPRU/Desktop/testdata.xlsx";
        FileInputStream fs = new FileInputStream(path);

        // Creating a workbook
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int row = sheet.getLastRowNum() + 1;

     // Open a single WebDriver instance outside the loop
        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, 10); // Wait up to 10 seconds

        for (int i = 1; i < row; i++) {
            driver.get("http://localhost:5173/addcourse");

            // Insert data to form fields
            String nameTH = sheet.getRow(i).getCell(1).toString();
            WebElement MajorNameTH = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[2]/form/div[1]/div[1]/input")));
            MajorNameTH.sendKeys(nameTH);
            
            String nameEng = sheet.getRow(i).getCell(2).toString();
            WebElement MajorNameEng = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[2]/form/div[1]/div[2]/input"));
            MajorNameEng.sendKeys(nameEng);
            
            String majorID = sheet.getRow(i).getCell(3).toString();
            WebElement inputMajorID = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[2]/form/div[2]/div[1]/input"));
            inputMajorID.sendKeys(majorID);
            
            String majorYear = sheet.getRow(i).getCell(4).toString();
            WebElement inputMajorYear = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[2]/form/div[2]/div[2]/input"));
            inputMajorYear.sendKeys(majorYear);
            
            String majorUnit = sheet.getRow(i).getCell(5).toString();
            WebElement inputMajorUnit = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[2]/form/div[2]/div[3]/input"));
            inputMajorUnit.sendKeys(majorUnit);

            // Click the submit button with wait
            WebElement submitButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[2]/form/div[4]/button[2]")));
            submitButton.click();

            // Validate result message
            String actual = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"my_modal_1\"]/div/h3"))).getText();
            String expected = sheet.getRow(i).getCell(6).toString();

            Row rows = sheet.getRow(i);
            Cell cell = rows.createCell(7);
            cell.setCellValue(actual);

            if (expected.equals(actual)) {
                Cell cell2 = rows.createCell(8);
                cell2.setCellValue("Pass");
            } else {
                Cell cell2 = rows.createCell(8);
                cell2.setCellValue("Fail");
            }

            Cell cell3 = rows.createCell(9);
            cell3.setCellValue(testDate);
            Cell cell4 = rows.createCell(10);
            cell4.setCellValue(testerName);
        }

        // Save the Excel file after the loop
        FileOutputStream fos = new FileOutputStream(path);
        workbook.write(fos);
        fos.close();

        driver.quit();
        workbook.close();
        fs.close();
    }
}
