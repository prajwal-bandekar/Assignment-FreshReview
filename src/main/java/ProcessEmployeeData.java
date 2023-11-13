import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

public class ProcessEmployeeData {
    public static void main(String[] args) {
        try {
            // Remove Duplicate Records
            String inputFilePath = "src/dummyData.xlsx";
            String cleanedDataFilePath = "src/cleanedDATA.xlsx";
            removeDuplicates(inputFilePath, cleanedDataFilePath);

            //Insert the got Unique Data into Database with generating unique Id's too
            String excelFilePath = "src/cleanedDATA.xlsx";
            insertIntoFinanceDB(excelFilePath);

        } catch (IOException | SQLException e) {
            e.printStackTrace();
        }
    }

    private static void removeDuplicates(String inputFilePath, String outputFilePath) throws IOException {
        try (FileInputStream inputStream = new FileInputStream(inputFilePath)) {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            Set<String> uniquePhoneNumbers = new HashSet<>();
            Iterator<Row> iterator = sheet.iterator();

            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Cell phoneCell = currentRow.getCell(2);
                String phoneNumber = phoneCell.getStringCellValue();

                if (uniquePhoneNumbers.contains(phoneNumber)) {
                    iterator.remove();
                } else {
                    uniquePhoneNumbers.add(phoneNumber);
                }
            }

            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("UniqueData");

            for (Row row : sheet) {
                Row newRow = newSheet.createRow(newSheet.getLastRowNum() + 1);

                for (Cell cell : row) {
                    Cell newCell = newRow.createCell(cell.getColumnIndex());
                    setCellValue(newCell, cell);
                }
            }

            try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                newWorkbook.write(outputStream);
                System.out.println("Distinct data saved to: " + outputFilePath);
            }
        }
    }

    private static void insertIntoFinanceDB(String excelFilePath) throws SQLException, IOException {
        try (FileInputStream excelInputStream = new FileInputStream(excelFilePath)) {
            Workbook excelWorkbook = new XSSFWorkbook(excelInputStream);
            Sheet excelSheet = excelWorkbook.getSheetAt(0);

            Connection connection = getDatabaseConnection();

            for (int i = 1; i <= excelSheet.getLastRowNum(); i++) {
                Row row = excelSheet.getRow(i);

                String firstName = getCellValueAsString(row.getCell(1));
                String lastName = getCellValueAsString(row.getCell(2));

                boolean employeeExists = checkIfEmployeeExists(connection, firstName, lastName);

                String uniqueID = generateUniqueID(connection, firstName, lastName, employeeExists);

                insertIntoFinanceDB(connection, row, uniqueID);
            }

            connection.close();
            System.out.println("Data insertion into Finance Database complete.");
        }
    }

    private static void setCellValue(Cell newCell, Cell oldCell) {
        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(oldCell)) {
                    newCell.setCellValue(oldCell.getDateCellValue());
                } else {
                    newCell.setCellValue(oldCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            default:
                newCell.setCellValue("");
        }
    }

    private static Connection getDatabaseConnection() throws SQLException {
        String url = "jdbc:mysql://localhost:3306/finance_database?sessionVariables=sql_mode='NO_ENGINE_SUBSTITUTION'&jdbcCompliantTruncation=false";
        String username = "root";
        String password = "openmysql";

        return DriverManager.getConnection(url, username, password);
    }

    private static boolean checkIfEmployeeExists(Connection connection, String firstName, String lastName) throws SQLException {
        String query = "SELECT COUNT(*) FROM employee_table WHERE first_name = ? AND last_name = ?";
        try (PreparedStatement preparedStatement = connection.prepareStatement(query)) {
            preparedStatement.setString(1, firstName);
            preparedStatement.setString(2, lastName);

            ResultSet resultSet = preparedStatement.executeQuery();
            resultSet.next();
            int count = resultSet.getInt(1);

            return count > 0;
        }
    }

    private static String generateUniqueID(Connection connection, String firstName, String lastName, boolean employeeExists) throws SQLException {
        String baseID = (firstName + lastName).toLowerCase();

        if (employeeExists) {
            int count = 1;
            String uniqueID = baseID + count;

            while (checkIfIDExists(connection, uniqueID)) {
                count++;
                uniqueID = baseID + count;
            }
            return uniqueID;
        } else {
            return baseID;
        }
    }

    private static boolean checkIfIDExists(Connection connection, String baseID) throws SQLException {
        String query = "SELECT COUNT(*) FROM employee_table WHERE unique_id = ?";
        try (PreparedStatement preparedStatement = connection.prepareStatement(query)) {
            preparedStatement.setString(1, baseID);

            ResultSet resultSet = preparedStatement.executeQuery();
            resultSet.next();
            int count = resultSet.getInt(1);

            return count > 0;
        }
    }

    private static void insertIntoFinanceDB(Connection connection, Row newRow, String uniqueID) throws SQLException {
        int serialNumber;
        try {
            serialNumber = (int) newRow.getCell(0).getNumericCellValue();
        } catch (IllegalStateException e) {
            System.out.println("Skipping row with non-numeric Serial Number");
            return;
        }

        String firstName = getCellValueAsString(newRow.getCell(1));
        String lastName = getCellValueAsString(newRow.getCell(2));
        int salary;
        try {
            salary = (int) newRow.getCell(3).getNumericCellValue();
        } catch (IllegalStateException e) {
            System.out.println("Skipping row with non-numeric Salary");
            return;
        }

        String jobPosition = getCellValueAsString(newRow.getCell(4));
        if (jobPosition.length() > 255) {
            jobPosition = jobPosition.substring(0, 255);
            System.out.println("Truncating job_position data");
        }

        int phoneNumber;
        try {
            phoneNumber = (int) newRow.getCell(5).getNumericCellValue();
        } catch (IllegalStateException e) {
        	//Handled an error when The Excel wasnot properly formatted
            System.out.println("Skipping row with non-numeric Phone Number");
            return;
        }

        String query = "INSERT INTO employee_table (serial_number, first_name, last_name, salary, job_position, unique_id, phone_number) VALUES (?, ?, ?, ?, ?, ?, ?)";
        try (PreparedStatement preparedStatement = connection.prepareStatement(query)) {
            preparedStatement.setInt(1, serialNumber);
            preparedStatement.setString(2, firstName);
            preparedStatement.setString(3, lastName);
            preparedStatement.setInt(4, salary);
            preparedStatement.setString(5, jobPosition);
            preparedStatement.setString(6, uniqueID);
            preparedStatement.setInt(7, phoneNumber);

            preparedStatement.executeUpdate();
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf((int) cell.getNumericCellValue());
        } else {
            return "";
        }
    }
}
