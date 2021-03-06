package Excel;
import org.apache.poi.ss.format.CellFormatType;
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
import java.lang.reflect.Type;
import java.util.Iterator;

public class MainClassOld {
    public static void main(String[] args) {
        ExcelParser.parse("объемы.xlsx");
    }
}

class ExcelParser1{
    public static void parse(String namefile) {
        try {
            File filenew = new File("D:/Моё развитие и дела/Программирование/Проекты/Excel/объемы.xlsx");
            FileInputStream file = new FileInputStream(filenew);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFSheet sheet2 = workbook.createSheet("List");
            Iterator<Row> rowIterator = sheet.iterator();
            int rowi=0;
            while (rowIterator.hasNext()) {
                Row row2= sheet2.createRow(rowi);
                rowi++;
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                int result=0;
                while (cellIterator.hasNext()) {


                    Cell cell = cellIterator.next();
                    Cell cell2= row2.createCell(result);
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue()+" ");
                            cell2=row2.createCell(result);
                            cell2.setCellValue(cell.getNumericCellValue());
                            result+=1;
                            break;
                        case Cell.CELL_TYPE_STRING:
                            String s = cell.getStringCellValue();
                            if(s.equals("карниз для штор")){
                                System.out.print(s+" ");
                                cell2=row2.createCell(result);
                                cell2.setCellValue(s);
                                cellIterator.next();
                                System.out.print("Демонтаж/Монтаж карниза для штор ");
                                cell2=row2.createCell(result);
                                cell2.setCellValue("Демонтаж/Монтаж карниза для штор");
                                result+=1;
                                continue;
                            }
                            if(s.equals("люстра")){
                                System.out.print(s+" ");
                                cell2=row2.createCell(result);
                                cell2.setCellValue(s);
                                cellIterator.next();
                                System.out.print("Демонтаж/Монтаж люстры ");
                                cell2=row2.createCell(result);
                                cell2.setCellValue("Демонтаж/Монтаж люстры");
                                result+=1;
                                continue;
                            }

                            s=s.replace("Шпатлевка потолчка","Шпатлевка потолка").replace(", Штукатурка потолка","").replace
                                    ("кухонный гарнитур (подвесная и напольная части)", "кухонный гарнитур").replace
                                    ("Монтаж кухонного гарнитура (подвесная часть), Демонтаж кухонного гарнитура (подвесная часть)", "Монтаж кухонного гарнитура, Демонтаж кухонного гарнитура").replace
                                    ("выключатель", "выключатель/розетка").replace("Монтаж выключателя, Демонтаж выключателя", "Монтаж выключателя/розетки, Демонтаж выключателя/розетки").replace
                                    ("Демонтаж выключателя, Монтаж выключателя","Монтаж выключателя/розетки, Демонтаж выключателя/розетки");
                            cell2=row2.createCell(result);
                            cell2.setCellValue(s);
                            System.out.print(s+" ");
                            result+=1;
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            System.out.print(cell.getNumericCellValue()+" ");
                            cell2=row2.createCell(result);
                            cell2.setCellValue(cell.getNumericCellValue());
                            result+=1;
                            break;
                    }

                }
                System.out.println("");
            }
            FileOutputStream file2=new FileOutputStream(new File("D:/Моё развитие и дела/Программирование/Проекты/Excel/объемы_новые.xlsx"));
            workbook.write(file2);
            file2.close();
            file.close();

        } catch(Exception e){
            e.printStackTrace();
        }
    }
}