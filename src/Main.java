import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.Scanner;

public class Main {

    public static final String consoleCyan = "\u001B[36m";
    public static final String consoleYellow = "\u001B[33m";
    public static final String consoleGreen = "\u001B[32m";
    public static final String consoleReset = "\u001B[0m";


    public static void main(String[] args) {

        Scanner scanner = new Scanner(System.in);
        System.out.println("EXCEL TO JSON CONVERTER (using Apache POI)");
        System.out.print(consoleCyan + "Enter Excel Input File Directory : " + consoleReset);
        String excelFilePath = scanner.nextLine();

        System.out.print(consoleCyan + "Enter Json Output File Directory : " + consoleReset);
        String jsonFilePath = scanner.nextLine();

        System.out.print(consoleCyan + "Enter Array Name : " + consoleReset);
        String arrayName = scanner.nextLine();

        String[] headerProperties;
        String[][] cellValues;

        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            FileOutputStream fos = new FileOutputStream(jsonFilePath);
            BufferedWriter bufferedWriter = new BufferedWriter(new OutputStreamWriter(fos));

            Workbook workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheetAt(0);

            Row headerRow = sheet.getRow(0);
            int lastCellNum = headerRow.getLastCellNum();

            headerProperties = new String[headerRow.getLastCellNum()];
            cellValues = new String[sheet.getLastRowNum()][headerRow.getLastCellNum()];

            for (int j = 0; j < lastCellNum; j++) {
                headerProperties[j] = headerRow.getCell(j).getStringCellValue();
            }

            System.out.println("[---- " + consoleYellow + "Reading Excel File..." + consoleReset + " ----]");
            System.out.println("> Starting...");

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row currentRow = sheet.getRow(i);
                for (int j = 0; j < lastCellNum; j++) {
                    Cell currentCell = currentRow.getCell(j);
                    String value = currentCell.getStringCellValue().replace("\"", "\\\"");
                    cellValues[i - 1][j] = value;
                }
                System.out.println("\r> " + i + " out of " + sheet.getLastRowNum() + " rows completed.");
            }
            workbook.close();
            fis.close();

            System.out.println(consoleGreen + "\r> Completed!" + consoleReset);
            System.out.println("[---- " + consoleYellow + "Writing Json File..." + consoleReset + " ----]");
            System.out.println("> Starting...");

            bufferedWriter.write("{\n\t\"" + arrayName + "\": [\n");
            for (int i = 0; i < cellValues.length; i++) {
                bufferedWriter.write("\t\t{\n");
                for (int j = 0; j < cellValues[i].length; j++) {
                    bufferedWriter.write("\t\t\t\"" + headerProperties[j] + "\": \"" + cellValues[i][j] + "\"");
                    if (j < cellValues[i].length - 1) {
                        bufferedWriter.write(",\n");
                    } else {
                        bufferedWriter.write("\n");
                    }
                    System.out.println("\r> " + i + " out of " + sheet.getLastRowNum() + " rows completed.");
                }
                bufferedWriter.write("\t\t}");
                if (i < cellValues.length - 1) {
                    bufferedWriter.write(",\n");
                } else {
                    bufferedWriter.write("\n");
                }
            }
            bufferedWriter.write("\t]\n}");

            System.out.println(consoleGreen + "\r> Completed!" + consoleReset);

            bufferedWriter.flush();
            bufferedWriter.close();

        } catch (Exception ex) {
            System.err.println("Exception occurred. Please check file.");
            ex.printStackTrace();
        }

    }
}