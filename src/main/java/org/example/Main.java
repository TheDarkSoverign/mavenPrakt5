package org.example;

import java.sql.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    protected static Scanner sc = new Scanner(System.in);
    protected static Connection con;
    static final String schema = "task5";
    protected static String table;

    static String Url = "jdbc:postgresql://localhost:5432/postgres";

    static {
        try {
            con = DriverManager.getConnection(Url, "postgres", "postgres");
        } catch (SQLException e) {
            System.out.println("Не удалось подключиться к базе данных: " + e.getMessage());
        }

        try {
            con.setAutoCommit(false);

            Statement st = con.createStatement();
            st.executeUpdate("CREATE SCHEMA IF NOT EXISTS " + schema);
            st.executeUpdate("SET search_path TO " + schema);

            con.commit();
            con.setAutoCommit(true);
            System.out.println("Используется схема - " + schema);
        } catch (SQLException e) {
            System.out.println("Не удалось создать схему для задания: " + e.getMessage());
            try {
                con.rollback();
            } catch (SQLException ex) {
                throw new RuntimeException(ex);
            }
        }

        String query = "CREATE TABLE IF NOT EXISTS task5 (id SERIAL, str1 TEXT, str2 TEXT, rev1 TEXT, rev2 TEXT, concat TEXT)";
        try {
            Statement st = con.createStatement();
            st.executeUpdate(query);
            System.out.println("Используется таблица по умолчанию - " + table);
        } catch (SQLException e) {
            System.out.println("Не удалось использовать таблицу по умолчанию, " + e.getMessage());
        }
    }

    protected static void menu() {
        int x = 0;
        String s = "";
        Task tasks = new Task();
        ExportToExcel export = new ExportToExcel();
        while (!"0".equals(s)) {
            System.out.println("Меню программы:");
            System.out.println("1. Вывести все таблицы.");
            System.out.println("2. Создать/выбрать таблицу.");
            System.out.println("3. Ввести две строки (минимум 50 символов).");
            System.out.println("4. Поменять порядок символов строк на обратный.");
            System.out.println("5. Соединить строки.");
            System.out.println("6. Записать результаты в таблицу и вывести.");
            System.out.println("7. Записать данные в Excel и вывести.");
            System.out.println("0. Выход");
            System.out.print("Выберите пункт меню: ");
            s = sc.nextLine();
            try {
                x = Integer.parseInt(s);
            } catch (NumberFormatException e) {
                System.out.println("Неверный формат ввода");
            }
            switch (x) {
                case 1 -> tasks.task1();
                case 2 -> tasks.task2();
                case 3 -> tasks.task3();
                case 4 -> tasks.task4();
                case 5 -> tasks.task5();
                case 6 -> {
                    tasks.insertData();
                    tasks.selectData();
                }
                case 7 -> {
                    System.out.print("Введите название файла: ");
                    String filepath = sc.nextLine();

                    if (!filepath.contains(".xlsx")) {
                        filepath += ".xlsx";
                    }

                    export.exportData(table, filepath);
                    export.printExcelData(filepath);
                }
                case 0 -> System.out.println("Пока!");
                default -> System.out.println("Неправильно выбран пункт меню! Попробуйте еще раз...");
            }
            x = 0;
        }
    }

    public static void main(String[] args) {
        System.out.println("Подключились к БД. ");
        menu();
    }
}

class Task extends Main {
    static StringBuffer sb1 = new StringBuffer();
    static StringBuffer sb2 = new StringBuffer();

    static StringBuffer rev1 = new StringBuffer();
    static StringBuffer rev2 = new StringBuffer();

    static StringBuffer concat = new StringBuffer();

    public void task1() {
        String query = "SELECT table_name AS Названия_таблиц FROM information_schema.tables WHERE table_schema = '" + schema + "'";
        try {
            Statement st = con.createStatement();
            ResultSet rs = st.executeQuery(query);
            try {
                int nameLength = 15;
                while (rs.next()) {
                    int currentNameLength = rs.getString(1).length();
                    if (currentNameLength > nameLength) {
                        nameLength = currentNameLength;
                    }
                }
                String tablePart = "+" + "-".repeat(5) + "+" + "-".repeat(nameLength + 2) + "+";
                System.out.println("Список таблиц:");
                System.out.println(tablePart);
                System.out.printf("| %-3s | %-" + nameLength + "s |\n", "ID", "Названия таблиц");
                rs = st.executeQuery(query);
                int i = 1;
                while (rs.next()) {
                    String tableName = rs.getString(1);
                    System.out.println(tablePart);
                    System.out.printf("| %-3d | %-" + nameLength + "s |\n", i++, tableName);
                }
                System.out.println(tablePart);
            } catch (SQLException e) {
                System.out.println("Не удалось вывести результат, " + e.getMessage());
            }
        } catch (SQLException e) {
            System.out.println("Не удалось выполнить запрос, " + e.getMessage());
        }
    }

    public void task2() {
        System.out.print("Введите название таблицы: ");
        table = sc.next();
        sc.nextLine();
        String query = "CREATE TABLE IF NOT EXISTS " + table + " (id SERIAL, str1 VARCHAR(255), str2 VARCHAR(255), rev1 VARCHAR(255), rev2 VARCHAR(255), concat VARCHAR(255))";
        try {
            Statement st = con.createStatement();
            st.executeUpdate(query);
            System.out.println("Таблица " + table + " успешно создана/выбрана!");
        } catch (SQLException e) {
            System.out.println("Не удалось выполнить запрос, " + e.getMessage());
            task2();
        }
    }

    public void task3() {
        System.out.println(inputFirstStr().toString());
        System.out.println("Сохранили строку.");
        System.out.println(inputSecondStr().toString());
        System.out.println("Сохранили строку.");
    }

    public void task4() {
        rev1 = new StringBuffer(sb1.reverse());
        System.out.println("Перевернутая первая строка:");
        System.out.println(rev1.toString());
        sb1.reverse();

        rev2 = new StringBuffer(sb2.reverse());
        System.out.println("Перевернутая вторая строка:");
        System.out.println(rev2.toString());
        sb2.reverse();
    }

    public void task5() {
        concat = new StringBuffer(sb1.append(sb2));
        System.out.println("Соединенная строка:");
        System.out.println(concat.toString());
        sb1.delete(sb2.length(), sb1.length());
    }

    public void insertData() {
        String query = "INSERT INTO " + table + " (str1, str2, rev1, rev2, concat) VALUES (?, ?, ?, ?, ?)";
        try (PreparedStatement pst = con.prepareStatement(query)) {
            pst.setObject(1, sb1.toString());
            pst.setObject(2, sb2.toString());
            pst.setObject(3, rev1.toString());
            pst.setObject(4, rev2.toString());
            pst.setObject(5, concat.toString());
            pst.executeUpdate();
            System.out.println("Все выполненные результаты добавлены в таблицу!");
        } catch (
                SQLException e) {
            System.out.println("Не удалось выполнить запрос, " + e.getMessage());
        }
    }

    public void selectData() {
        System.out.println("Получаю данные...");
        String query = "SELECT * FROM " + table;
        try (PreparedStatement pst = con.prepareStatement(query);
             ResultSet rs = pst.executeQuery()) {

            ResultSetMetaData rsmd = rs.getMetaData();
            int length = rsmd.getColumnCount();

            String[] columnNames = new String[length];
            int[] maxLength = new int[length];

            List<List<String>> rows = new ArrayList<>();

            for (int i = 0; i < length; i++) {
                columnNames[i] = rsmd.getColumnName(i + 1);
                maxLength[i] = rsmd.getColumnName(i + 1).length();
            }

            while (rs.next()) {
                List<String> row = new ArrayList<>();
                for (int i = 0; i < length; i++) {
                    String obj = rs.getString(i + 1);
                    obj = (obj != null) ? obj : "NULL";
                    row.add(obj);
                    if (obj.length() > maxLength[i]) {
                        maxLength[i] = obj.length();
                    }
                }
                rows.add(row);
            }

            StringBuilder border = new StringBuilder("+");
            StringBuilder header = new StringBuilder("|");

            for (int width : maxLength) {
                border.append("-".repeat(width + 2)).append("+");
            }
            System.out.println("Полученные данные из таблицы: ");
            System.out.println(border);


            for (int i = 0; i < length; i++) {
                header.append(" ").append(String.format("%-" + maxLength[i] + "s", columnNames[i])).append(" |");
            }
            System.out.println(header);
            System.out.println(border);

            for (List<String> row : rows) {
                StringBuilder rowStr = new StringBuilder("|");
                for (int i = 0; i < length; i++) {
                    String val = (i < row.size()) ? row.get(i) : "";
                    rowStr.append(" ").append(String.format("%-" + maxLength[i] + "s", val)).append(" |");
                }
                System.out.println(rowStr);
                System.out.println(border);
            }

        } catch (SQLException e) {
            System.out.println("Не удалось получить данные из таблицы, " + e.getMessage());
        }
    }

    public StringBuffer inputFirstStr() {
        sb1.delete(0, sb1.length());
        System.out.println("Введите первую строку!");
        return getStringBuffer(sb1);
    }


    public StringBuffer inputSecondStr() {
        sb2.delete(0, sb2.length());
        System.out.println("Введите вторую строку!");
        return getStringBuffer(sb2);
    }

    private StringBuffer getStringBuffer(StringBuffer sb) {
        while (sb.length() < 50) {
            sb.delete(0, sb.length());
            System.out.println("v".repeat(50));
            if (sb.append(sc.nextLine()).length() < 50) {
                System.out.println("Длина строки меньше 50 символов!");
            }
        }
        return sb;
    }
}

class ExportToExcel extends Main {

    public void exportData(String table, String filepath) {
        String printAll = "SELECT * FROM " + table;
        try (PreparedStatement pst = con.prepareStatement(printAll); ResultSet rs = pst.executeQuery()) {
            Workbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet("task 1");
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue(rs.getMetaData().getColumnName(1));
            row.createCell(1).setCellValue(rs.getMetaData().getColumnName(2));
            row.createCell(2).setCellValue(rs.getMetaData().getColumnName(3));
            row.createCell(3).setCellValue(rs.getMetaData().getColumnName(4));
            row.createCell(4).setCellValue(rs.getMetaData().getColumnName(5));
            row.createCell(5).setCellValue(rs.getMetaData().getColumnName(6));

            int rowIndex = 1;
            while (rs.next()) {
                Row row1 = sheet.createRow(rowIndex++);
                row1.createCell(0).setCellValue(rs.getInt(1));
                row1.createCell(1).setCellValue(rs.getString(2));
                row1.createCell(2).setCellValue(rs.getString(3));
                row1.createCell(3).setCellValue(rs.getString(4));
                row1.createCell(4).setCellValue(rs.getString(5));
                row1.createCell(5).setCellValue(rs.getString(6));

            }
            int columnCount = sheet.getRow(0).getPhysicalNumberOfCells();
            for (int i = 0; i < columnCount; i++) {
                sheet.autoSizeColumn(i);
            }
            try (FileOutputStream fos = new FileOutputStream(filepath)) {
                wb.write(fos);
            } catch (IOException e) {
                System.out.println("Ошибка при записи Excel-файла: " + e);
            } finally {
                wb.close();
                System.out.println("Данные успешно сохранены в Excel-файл: " + filepath);
            }
        } catch (IOException | SQLException e) {
            System.out.println("Ошибка при экспорте данных: " + e);
        }
    }

    public void printExcelData(String filepath) {
        try (Workbook wb = new XSSFWorkbook(filepath)) {
            Sheet sheet = wb.getSheetAt(0);
            System.out.println("\nДанные из Excel:");


            StringBuilder border = new StringBuilder("+");

            Row names = sheet.getRow(0);

            int[] maxLength = new int[names.getLastCellNum()];

            for (Cell name : names) {
                maxLength[name.getColumnIndex()] = name.toString().length();

            }
            for (Row row : sheet) {
                for (int i = 0; i < maxLength.length; i++) {
                    Cell obj = row.getCell(i);
                    int length = obj.toString().length();
                    if (length > maxLength[i]) {
                        maxLength[i] = length;
                    }
                }
            }

            for (int width : maxLength) {
                border.append("-".repeat(width + 2)).append("+");
            }

            System.out.println(border);

            for (Row row : sheet) {
                StringBuilder rowStr = new StringBuilder("|");
                for (int i = 0; i < maxLength.length; i++) {
                    Cell obj = row.getCell(i);
                    rowStr.append(" ").append(String.format("%-" + maxLength[i] + "s", obj)).append(" |");
                }
                System.out.println(rowStr);
                System.out.println(border);
            }

        } catch (IOException e) {
            System.out.println("Ошибка при чтении Excel-файла: " + e.getMessage());
        }
    }
}
