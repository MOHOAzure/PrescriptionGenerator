package hby;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

import medic.Prescription;

public class ExcelReader {
  private ArrayList < Prescription > prescriptionList;
  private String departDate;

  public ExcelReader() {
    this.prescriptionList = new ArrayList < > ();
    this.departDate = "";
  }

  public ArrayList < Prescription > readPrescriptionExcel(String file, String date) {
    HSSFWorkbook workbook = null;
    FileInputStream excelFile = null;
    HSSFSheet sheet = null;
    int rowNum = 0, progress = 0;
    String houseName, residentName, departmentName;

    try {
      excelFile = new FileInputStream(new File(file).getCanonicalPath());
      workbook = new HSSFWorkbook(excelFile);
      sheet = workbook.getSheet(date);
      Iterator < Row > rowIterator = sheet.iterator();

      while (rowIterator.hasNext()) {
        Row row = rowIterator.next();

        //get departDate
        if (rowNum == 1) {
          setDepartDate(row.getCell(0).getStringCellValue().trim());
        }

        //get prescriptions
        if (rowNum >= 3) {
          houseName = row.getCell(1).getStringCellValue().trim();

          if (houseName.isEmpty())
            continue;

          residentName = row.getCell(2).getStringCellValue().trim();
          departmentName = row.getCell(3).getStringCellValue().trim();
          progress = (int) row.getCell(5).getNumericCellValue();
          /**/
          System.out.println("戶別 : " + row.getCell(1).getStringCellValue().trim());
          System.out.println("住民 : " + row.getCell(2).getStringCellValue().trim());
          System.out.println("科別 : " + row.getCell(3).getStringCellValue().trim());
          //System.out.println("診號 : "+row.getCell(4).getStringCellValue().trim());
          System.out.println("進度 : " + (int) row.getCell(5).getNumericCellValue());
          //System.out.println("醫生 : "+row.getCell(6).getStringCellValue().trim());
          System.out.println("");

          Prescription prescrip = new Prescription(
            houseName,
            residentName,
            departmentName,
            progress);
          this.prescriptionList.add(prescrip);
        }
        rowNum++;
      }
      //this.prescriptionList.forEach(prescrip -> System.out.println(prescrip));
      //System.out.println("Prescription Excel file is loaded");
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    } finally {
      if (null != workbook) {
        try {
          workbook.close();
        } catch (IOException e) {
          e.printStackTrace();
        }
      }
      if (null != excelFile) {
        try {
          excelFile.close();
        } catch (IOException e) {
          e.printStackTrace();
        }
      }
    }
    return this.prescriptionList;
  }

  public String getDepartDate() {
    return departDate;
  }

  public void setDepartDate(String departDate) {
    this.departDate = departDate;
  }

  public void getSheetNames(String file) {
    HSSFWorkbook workbook = null;
    FileInputStream excelFile = null;

    try {
      excelFile = new FileInputStream(new File(file).getCanonicalPath());
      workbook = new HSSFWorkbook(excelFile);

      // for each sheet in the workbook
      for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
        System.out.println("Sheet name: " + workbook.getSheetName(i));
      }

    } catch (IOException e) {
      e.printStackTrace();
    } finally {
      if (excelFile != null) {
        try {
          excelFile.close();
        } catch (IOException e) {
          e.printStackTrace();
        }
      }
    }
  }
}
