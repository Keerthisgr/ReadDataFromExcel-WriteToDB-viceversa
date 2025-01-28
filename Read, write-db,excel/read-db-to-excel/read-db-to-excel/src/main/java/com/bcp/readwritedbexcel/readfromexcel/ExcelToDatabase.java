package com.bcp.readwritedbexcel.readfromexcel;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;

public class ExcelToDatabase {
    public static void main(String[] args) throws SQLException, IOException {
        // Connect to the database
        String URL = "jdbc:mysql://localhost:3306/rw_db_excel";
        String USER = "root";
        String PASS = "Keerthi8088169847";
        Connection conn = DriverManager.getConnection(URL, USER, PASS);
        Statement stmt = conn.createStatement();

        // Create a new table in the database
        String sql = "CREATE TABLE IF NOT EXISTS Restaurant (" +
                "r_id INT, " +
                "r_open TINYINT(1), " +
                "r_location VARCHAR(50), " +
                "r_ratings DOUBLE, " +
                "r_name VARCHAR(25))";
        stmt.execute(sql);

        // Read Excel file
        FileInputStream fis = new FileInputStream("C:\\Users\\keert\\Desktop\\Read, write-db,excel\\read-db-to-excel\\read-db-to-excel\\datafiles\\restro.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Restro Data");
        int lastRowNum = sheet.getLastRowNum();

        // Insert data into the database
        for (int r = 1; r <= lastRowNum; r++) {
            XSSFRow row = sheet.getRow(r);

            int resId = (int) row.getCell(0).getNumericCellValue();
            int resOpen = row.getCell(1).getBooleanCellValue() ? 1 : 0; // Convert boolean to 1/0
            String resLoc = row.getCell(2).getStringCellValue();
            double resRating = row.getCell(3).getNumericCellValue();
            String resName = row.getCell(4).getStringCellValue();

            // Create the SQL insert statement
            sql = "INSERT INTO Restaurant VALUES (" +
                    resId + ", " +
                    resOpen + ", '" +
                    resLoc + "', " +
                    resRating + ", '" +
                    resName + "')";
            stmt.execute(sql);
        }

        workbook.close();
        fis.close();
        conn.close();

        System.out.println("Successfully read data from Excel and wrote it into the database table.");
    }
}
