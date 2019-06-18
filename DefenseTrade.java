package com.aml.excel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class DefenseTrade {

  // Chromedriver path
    static {
        System.setProperty("webdriver.chrome.driver", "/home/obsessory/Documents/raaki/proj/soft/chromedriver_linux64/chromedriver");
    }

    static List<DefenseTradeUSModel> DefenseTradeList= new ArrayList<DefenseTradeUSModel>();

    public static void main(String[] args) throws IOException {
        String downloadPath = "/home/obsessory/Documents/KYC";
        Map<String, Object> preferences = new Hashtable<String, Object>();
        preferences.put("profile.default_content_settings.popups", 0);
        preferences.put("download.prompt_for_download", "false");
        preferences.put("download.default_directory", downloadPath);

        ChromeOptions options = new ChromeOptions();
        options.setExperimentalOption("prefs", preferences);
        DesiredCapabilities capabilities = DesiredCapabilities.chrome();
        capabilities.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
        capabilities.setCapability(ChromeOptions.CAPABILITY, options);

        File src = new File("/home/obsessory/Documents/KYC/Stat Debarred Parties_5.10.18.xlsx");
        src.delete();
        WebDriver driver = new ChromeDriver(capabilities);
        driver.get("https://www.pmddtc.state.gov/?id=ddtc_kb_article_page&sys_id=c22d1833dbb8d300d0a370131f9619f0");
        driver.findElement(By.xpath("//*[@id='maincontent']/div/div/div[2]/div/div/ul[1]/li[2]/a")).click();
        while (true) {
            if (src.exists()) {
                // System.out.println("Downloaded");
                break;
            } else {
                continue;
            }
        }
        driver.close();

        FileInputStream fis = new FileInputStream(src);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        int rownum = sheet.getLastRowNum();



        for (int i = 1; i <= rownum; i++) {

            String partyName = sheet.getRow(i).getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue();
            String  dob= sheet.getRow(i).getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue();
            String federalRegisterNotice = sheet.getRow(i).getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue();
            Date noticeDate = sheet.getRow(i).getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getDateCellValue();
            String correctedNotice = sheet.getRow(i).getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue();
            Date correctedNoticedate = sheet.getRow(i).getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getDateCellValue();
            DefenseTradeUSModel defenseTrade = new DefenseTradeUSModel();
            defenseTrade.setParty_Name(partyName);
            defenseTrade.setDate_Of_Birth(dob);
            defenseTrade.setFederal_Register_Notice(federalRegisterNotice);
            defenseTrade.setNotice_Date(noticeDate);
            defenseTrade.setCorrected_Notice(correctedNotice);
            defenseTrade.setCorrected_Notice_date(correctedNoticedate);

            DefenseTradeList.add(defenseTrade);
            System.out.println(defenseTrade.getParty_Name());



    }




    }


}
