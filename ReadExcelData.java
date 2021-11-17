package utils;

import org.apache.poi.xssf.usermodel.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.io.*;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashMap;

public class ReadExcelData
{
    public static void main(String args[]) throws IOException, JSONException
    {
        ReadExcelData read = new ReadExcelData();
        read.getRowCount();
    }

    private void getRowCount() throws IOException, JSONException
    {

        String excelFilePath = "C:\\Users\\prasanna.prabakaran\\Downloads\\MarkSheet.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        FileWriter file = new FileWriter("C:\\Users\\prasanna.prabakaran\\IdeaProjects\\Read excel and convert to JSON\\src\\output1.json");
        JSONObject sutdent=new JSONObject();
        BufferedWriter bufferedWriter =new BufferedWriter(file);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
       String name= String.valueOf(workbook.getSheetAt(0).getRow(0).getCell(0));
       String age= String.valueOf(workbook.getSheetAt(0).getRow(0).getCell(1));
       String marks= String.valueOf(workbook.getSheetAt(0).getRow(0).getCell(2));

        // Using FOR loop

        int lengthOfRows = sheet.getLastRowNum();
        int lengthOfCells = sheet.getRow(1).getLastCellNum();

       // LinkedHashMap<String, XSSFCell> data =new LinkedHashMap<>();
        JSONArray jsonArray =new JSONArray();

        for (int rowIndex = 1; rowIndex < lengthOfRows; rowIndex++)
        {
            JSONObject nameObj = new JSONObject();
            JSONObject ageObj = new JSONObject();
            JSONObject marksObj = new JSONObject();
            JSONArray students =new JSONArray();
            XSSFRow rowObj = sheet.getRow(rowIndex);
            for (int cellIndex = 0; cellIndex < lengthOfCells; cellIndex++)
            {

                XSSFCell cellObj = rowObj.getCell(cellIndex);
                //System.out.println(cellObj);
                switch (cellIndex){
                    case 0:
                        nameObj.put(name,cellObj);
                        students.put(nameObj);
                        break;
                    case 1:
                        ageObj.put(age,cellObj);
                        students.put(ageObj);
                        break;
                    case 2:
                        marksObj.put(marks,cellObj);
                        students.put(marksObj);
                        break;

                }
            }
            jsonArray.put(students);
            //System.out.println(jsonObject);

           // System.out.println();
        }
        bufferedWriter.write(String.valueOf(jsonArray));
        bufferedWriter.flush();
        bufferedWriter.newLine();


        file.close();
    }
}

     /* Iterator iterator= sheet.iterator();
      while (iterator.hasNext()){
         XSSFRow row= (XSSFRow) iterator.next();
        Iterator cellIterator= row.cellIterator();
        while (cellIterator.hasNext()){
           XSSFCell cell= (XSSFCell) cellIterator.next();
            switch (cell.getCellType()){
                case STRING:
                    //System.out.println("1");
                    System.out.print(cell.getStringCellValue());
                    break;

                case NUMERIC:
                    System.out.print(cell.getNumericCellValue());
                    break;
            }
            System.out.print("   ");


        }
          System.out.println();

      }*/



