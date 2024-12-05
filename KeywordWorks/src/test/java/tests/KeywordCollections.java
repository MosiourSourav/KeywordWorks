package tests;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class KeywordCollections {

    public static void main(String[] args) throws IOException, InterruptedException {
        WebDriver driver = new EdgeDriver();
        driver.manage().window().maximize();
        driver.get("https://www.google.com");

        String excelFilePath = "E:\\Keywords.xlsx";
        String[] sheetNames = {"Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"};

        for(String sheet : sheetNames) {
            List<String> values = getSheetKeywords(excelFilePath, sheet);

            int rowIndex = 2;
            for (String value : values) {
                driver.findElement(By.cssSelector("[name='q']")).sendKeys(value);
                Thread.sleep(5000);
                List<WebElement> options = driver.findElements(By.xpath("//ul[@role='listbox']/li//div[@role='option']/div[1]/span"));

                List<String> optionsList = new ArrayList<>();
                for (WebElement option : options) {
                    String text = option.getText();
                    optionsList.add(text);
                }

                List<String> longAndShortOptions = new ArrayList<>();
                longAndShortOptions.add(getLongestString(optionsList));
                longAndShortOptions.add(getShortestString(optionsList));

                writeIntoSpecificCells(excelFilePath, sheet, longAndShortOptions, rowIndex);
                driver.findElement(By.cssSelector("[name='q']")).clear();
                rowIndex++;
            }
        }

        driver.close();
    }

    public static List<String> getSheetKeywords(String excelFilePath, String sheetName) throws IOException {
        List<String> values = new ArrayList<>();
        String value = null;
        int columnIndex = 2;

        try {
            FileInputStream fis = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);

            for (Row row : sheet) {
                Cell cell = row.getCell(columnIndex);

                if (cell != null) {
                    value = cell.getStringCellValue();
                    values.add(value);
                }
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return values;
    }

    public static void writeIntoSpecificCells(String excelFilePath, String sheetName, List<String> longAndShortOptions, int rowIndex) {
        try {
            FileInputStream fis = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);

            Row row = sheet.getRow(rowIndex);

            int i = 3;
            for(String option : longAndShortOptions) {
                Cell cell = row.createCell(i);
                cell.setCellValue(option);
                i++;
            }
            fis.close();
            FileOutputStream fos = new FileOutputStream(excelFilePath);
            workbook.write(fos);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String getLongestString(List<String> optionsList) {
        String longest = optionsList.getFirst();

        for (String option : optionsList) {
            if (option.length() > longest.length()) {
                longest = option;
            }
        }
        return longest;
    }

    public static String getShortestString(List<String> optionsList) {
        String shortest = optionsList.getFirst();

        for (String option : optionsList) {
            if (option.length() < shortest.length()) {
                shortest = option;
            }
        }
        return shortest;
    }

}
