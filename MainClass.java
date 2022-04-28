package Excel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class MainClass {
    public static void main(String[] args) {
        ExcelParser.parse("объемы.xlsx");
    }
}

class ExcelParser{
    public static void parse(String namefile) {
        try {
            File filenew = new File("D:/Моё развитие и дела/Программирование/Проекты/Excel/объемы.xlsx");
            FileInputStream file = new FileInputStream(filenew);
            FileOutputStream file2= new FileOutputStream(filenew);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue()+" ");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            String s = cell.getStringCellValue();
                            if(s.equals("карниз для штор")){
                                System.out.print(s+" ");
                                cellIterator.next();
                                System.out.print("Демонтаж/Монтаж карниза для штор ");
                                continue;
                            }
                            if(s.equals("люстра")){
                                System.out.print(s+" ");
                                cellIterator.next();
                                System.out.print("Демонтаж/Монтаж люстры ");
                                continue;
                            }
                            s=s.replace("Шпатлевка потолчка","Шпатлевка потолка").replace(", Штукатурка потолка","").replace
                                    ("кухонный гарнитур (подвесная и напольная части)", "кухонный гарнитур").replace
                                    ("Монтаж кухонного гарнитура (подвесная часть), Демонтаж кухонного гарнитура (подвесная часть)", "Монтаж кухонного гарнитура, Демонтаж кухонного гарнитура").replace
                                    ("выключатель", "выключатель/розетка").replace("Монтаж выключателя, Демонтаж выключателя", "Монтаж выключателя/розетки, Демонтаж выключателя/розетки");

                            System.out.print(s+" ");
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            System.out.print(cell.getNumericCellValue()+" ");
                            break;
                    }
                }
                System.out.println("");
            }
            file.close();
        } catch(Exception e){
            e.printStackTrace();
        }
    }
}
