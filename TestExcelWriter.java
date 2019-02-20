package com.business.world.util;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

import com.business.world.entity.EmployeeEntity;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class TestExcelWriter {

	private static final String FILE_PATH = "/IdeaSourceCode/BusinessWorld/EmployeeList.xlsx";
	@Test
	public void testExcelWriter(){

	    List<EmployeeEntity> empList = new ArrayList<>();
        Workbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = (XSSFSheet) workbook.createSheet("Employee Data");


        //This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        int counter = 0;
        data.put(new Integer(counter++).toString(), new Object[] {"EMPLOYEE_ID","ID", "FIRSTNAME", "LASTNAME", "SALARY", "CITY", "STATE"});
        for(EmployeeEntity e : empList) {
            data.put(new Integer(counter++).toString(), new Object[] {e.getEmployeeId(), e.getId(), e.getFirstname(), e.getLastname(), e.getSalary(), "ABC1", "XYZ1"});
        }

       /* data.put("2", new Object[] {1, emp.getEmployeeId(), emp.getFirstname(), emp.getLastname(), 1234, "ABC1", "XYZ1"});
        data.put("3", new Object[] {2, "U60101", "Lokesh", "Gupta", 1234, "ABC2", "XYZ2"});
        data.put("4", new Object[] {3, "U60102", "John", "Adwards", 1234, "ABC3", "XYZ3"});
        data.put("5", new Object[] {4, "U60103", "Brian", "Schultz", 1234, "ABC4", "XYZ4"});*/
          
        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(FILE_PATH));
            workbook.write(out);
            out.close();
            System.out.println("EmployeeList.xlsx written successfully on disk within the project.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
