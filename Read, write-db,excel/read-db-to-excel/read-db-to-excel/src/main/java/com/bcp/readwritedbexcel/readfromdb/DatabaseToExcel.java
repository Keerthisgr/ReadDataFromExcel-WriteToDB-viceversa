package com.bcp.readwritedbexcel.readfromdb;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

public class DatabaseToExcel {
    public static void main(String[] args) throws SQLException, IOException {
        //connect to db
        String URL = "jdbc:mysql://localhost:3306/rw_db_excel";
        String USER = "root";
        String PASS = "Keerthi8088169847";
        Connection conn = DriverManager.getConnection(URL, USER, PASS);

        //stmt/query
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("SELECT * FROM restaurant");

        //excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Restro Data");

        XSSFRow row = sheet.createRow(0);
       row.createCell(0).setCellValue("restro_id");
       row.createCell(1).setCellValue("isOpen");
       row.createCell(2).setCellValue("location");
       row.createCell(3).setCellValue("rating");
       row.createCell(4).setCellValue("restroName");

       int r = 1;
       while (rs.next()){
           int restoId = rs.getInt("r_id");
           boolean isRestroOpem = rs.getBoolean("r_Open");
           String restroLocation = rs.getString("r_location");
           double restroRating = rs.getDouble("r_ratings");
           String nameOfRestro = rs.getString("r_name");

           row = sheet.createRow(r++);

           row.createCell(0).setCellValue(restoId);
           row.createCell(1).setCellValue(isRestroOpem);
           row.createCell(2).setCellValue(restroLocation);
           row.createCell(3).setCellValue(restroRating);
           row.createCell(4).setCellValue(nameOfRestro);

       }

       FileOutputStream fos = new FileOutputStream("C:\\Users\\keert\\Desktop\\Read, write-db,excel\\read-db-to-excel\\read-db-to-excel\\datafiles\\restroDetailss.xlsx");
       workbook.write(fos);
       workbook.close();
       fos.close();
       conn.close();

        System.out.println("Successfully Done");
    }
}
