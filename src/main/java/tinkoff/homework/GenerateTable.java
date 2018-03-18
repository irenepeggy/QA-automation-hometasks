package tinkoff.homework;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.logging.Logger;


public class GenerateTable {
    private static Logger log = Logger.getLogger(GenerateTable.class.getName());

    private void writeToTable(String filename) throws IOException {
        String shname = "Данные";
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(shname);

        int nrows = (int)(Math.random() * 30 + 1);
        makeHeaderLine(sheet);
        makeSheet(nrows, sheet, wb);
        for (int j = 0; j < sheet.getRow(0).getLastCellNum(); ++j) {
            sheet.autoSizeColumn(j);
        }
        wb.write(new FileOutputStream(filename));
        File file = new File(filename);
        String path = file.getAbsolutePath();

        log.info("Файл создан. Путь: " +  path);
        wb.close();
    }

    private void makeHeaderLine(XSSFSheet sheet){
        XSSFRow row = sheet.createRow(0);
        String[] head_cell_value = {"Имя", "Фамилия", "Отчество", "Возраст", "Пол",
                                    "Дата рождения", "Место рождения", "Почтовый индекс",
                                    "Страна", "Область", "Город", "Улица", "Дом", "Квартира"};
        for (int i = 0; i < head_cell_value.length; ++i) {
            row.createCell(i).setCellValue(head_cell_value[i]);
        }
    }

    private void makeSheet(int nrows, XSSFSheet sheet, XSSFWorkbook book) throws IOException {
        for (int i = 1; i <= nrows; i++) {
            XSSFRow row = sheet.createRow(i);
            writeToRow(row, book);
        }
    }

    private void writeToRow(XSSFRow row, XSSFWorkbook book) throws IOException {
        int sx = (int)(Math.random() * 2);
        ClassLoader classLoader;
        classLoader = getClass().getClassLoader();
        int j = 0;
        if (sx == 0){
            row.createCell(j++).setCellValue(getValue(classLoader.getResource("female_names.xlsx").getFile()));
            row.createCell(j++).setCellValue(getValue(classLoader.getResource("female_surnames.xlsx").getFile()));
            row.createCell(j++).setCellValue(getValue(classLoader.getResource("female_middlenames.xlsx").getFile()));
            row.createCell(4).setCellValue("Ж");
        } else {
            row.createCell(j++).setCellValue(getValue(classLoader.getResource("male_names.xlsx").getFile()));
            row.createCell(j++).setCellValue(getValue(classLoader.getResource("male_surnames.xlsx").getFile()));
            row.createCell(j++).setCellValue(getValue(classLoader.getResource("male_middlenames.xlsx").getFile()));
            row.createCell(4).setCellValue("М");
        }
        row.createCell(j).setCellValue((int)(Math.random() * 80 + 20)); // set age
        j += 2;
        DataFormat format = book.createDataFormat();
        CellStyle dateStyle = book.createCellStyle();
        dateStyle.setDataFormat(format.getFormat("dd-mm-yyyy"));
        row.createCell(j).setCellStyle(dateStyle);
        row.getCell(j++).setCellValue(new Date((long)(Math.random() * Integer.MAX_VALUE * 500 - Integer.MAX_VALUE * 5000 )));

        row.createCell(j++).setCellValue(getValue(classLoader.getResource("city.xlsx").getFile())); //set birth city
        row.createCell(j++).setCellValue((int)(Math.random() * 100000) + 100000); //set post index
        row.createCell(j++).setCellValue(getValue(classLoader.getResource("country.xlsx").getFile())); // set country
        row.createCell(j++).setCellValue(getValue(classLoader.getResource("region.xlsx").getFile())); // set region
        row.createCell(j++).setCellValue(getValue(classLoader.getResource("city.xlsx").getFile())); // set city
        row.createCell(j++).setCellValue(getValue(classLoader.getResource("street.xlsx").getFile())); // set street
        row.createCell(j++).setCellValue((int)(Math.random() * 100) + 1); // set house number
        row.createCell(j).setCellValue((int)(Math.random() * 200) + 1); // set room number*/
    }

    private  String getValue(String filename) throws IOException {
        XSSFWorkbook rsc = new XSSFWorkbook(new FileInputStream(filename));
        XSSFSheet sheet = rsc.getSheetAt(0);

        int rownum = (int)(Math.random() * sheet.getLastRowNum());
        XSSFRow row = sheet.getRow(rownum);
        return row.getCell(0).getStringCellValue();
    }

    public static void main(String[] args) throws IOException {
        String filename = "new_table.xlsx";
        GenerateTable generate = new GenerateTable();
        generate.writeToTable(filename);
    }
}
