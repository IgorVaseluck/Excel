package Excel;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Type;
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
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFSheet sheet2 = workbook.createSheet("List");
            XSSFCellStyle style= workbook.createCellStyle();
            style.setWrapText(true);
            Iterator<Row> rowIterator = sheet.iterator();
            int rowi=0;
            while (rowIterator.hasNext()) {
                Row row2= sheet2.createRow(rowi);
                row2.setHeightInPoints(45f);
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
                            cell2.setCellType(Cell.CELL_TYPE_NUMERIC);
                            cell2.setCellValue(cell.getNumericCellValue());
                            cell2.setCellStyle(style);
                            int windth=(int)(10*1.14388)*256;
                            sheet2.setColumnWidth(result,windth);
                            result+=1;
                            continue;
                        case Cell.CELL_TYPE_STRING:
                            String s = cell.getStringCellValue();
                            if (s.equals("кухня 7.3")||s.equals("Работы")||s.equals("Место")||s.equals("Наименование")||s.equals("Ед-цы измерения")||s.equals("Количество единиц")){
                                System.out.print(s+" ");
                                cell2=row2.createCell(result);
                                cell2.setCellType(Cell.CELL_TYPE_STRING);
                                cell2.setCellStyle(style);
                                int windth2=(int)(15*1.14388)*256;
                                if(s.equals("кухня 7.3")){cell2.setCellValue("кухня 7.3");}
                                if(s.equals("Работы")){cell2.setCellValue("Работы");}
                                if(s.equals("Место")){cell2.setCellValue("Место");}
                                if(s.equals("Наименование")){cell2.setCellValue("Наименование");}
                                if(s.equals("Ед-цы измерения")){cell2.setCellValue("Ед-цы измерения");}
                                if(s.equals("Количество единиц")){cell2.setCellValue("Количество единиц");}
                                sheet2.setColumnWidth(result,windth2);
                                row2.setHeightInPoints(15f);
                                result+=1;
                                continue;
                            }
                            if (s.equals("Помещение")){
                                System.out.print(s+" ");
                                cell2=row2.createCell(result);
                                cell2.setCellType(Cell.CELL_TYPE_STRING);
                                cell2.setCellStyle(style);
                                int windth2=(int)(15*1.14388)*256;
                                row2.setHeightInPoints(15f);
                                cell2.setCellValue("Помещение");
                                sheet2.setColumnWidth(result,windth2);
                                result+=1;
                                continue;
                            }
                            if (s.equals("шт.")){
                                System.out.print(s+" ");
                                cell2=row2.createCell(result);
                                cell2.setCellType(Cell.CELL_TYPE_STRING);
                                cell2.setCellStyle(style);
                                int windth2=(int)(15*1.14388)*256;
                                cell2.setCellValue("шт.");
                                sheet2.setColumnWidth(result,windth2);
                                result+=1;
                                continue;
                            }
                            if(s.equals("карниз для штор")){
                                System.out.print(s+" ");
                                cell2=row2.createCell(result);
                                cell2.setCellType(Cell.CELL_TYPE_STRING);
                                cell2.setCellStyle(style);
                                cell2.setCellValue(s);
                                int windth2=(int)(60*1.14388)*256;
                                sheet2.setColumnWidth(result,windth2);
                                result+=1;
                                cellIterator.next();
                                System.out.print("Демонтаж/Монтаж карниза для штор ");
                                cell2=row2.createCell(result);
                                cell2.setCellType(Cell.CELL_TYPE_STRING);
                                cell2.setCellStyle(style);
                                cell2.setCellValue("Демонтаж/Монтаж карниза для штор");
                                sheet2.setColumnWidth(result,windth2);
                                result+=1;
                                continue;
                            }
                            if(s.equals("люстра")){
                                System.out.print(s+" ");
                                cell2=row2.createCell(result);
                                cell2.setCellType(Cell.CELL_TYPE_STRING);
                                cell2.setCellStyle(style);
                                cell2.setCellValue(s);
                                int windth2=(int)(60*1.14388)*256;
                                sheet2.setColumnWidth(result,windth2);
                                result+=1;
                                cellIterator.next();
                                System.out.print("Демонтаж/Монтаж люстры ");
                                cell2=row2.createCell(result);
                                cell2.setCellType(Cell.CELL_TYPE_STRING);
                                cell2.setCellStyle(style);
                                cell2.setCellValue("Демонтаж/Монтаж люстры");
                                sheet2.setColumnWidth(result,windth2);
                                result+=1;
                                continue;
                            }

                            s=s.replace("Шпатлевка потолчка","Шпатлевка потолка").replace(", Штукатурка потолка","").replace
                                    ("кухонный гарнитур (подвесная и напольная части)", "кухонный гарнитур").replace
                                    ("Монтаж кухонного гарнитура (подвесная часть), Демонтаж кухонного гарнитура (подвесная часть)", "Монтаж кухонного гарнитура, Демонтаж кухонного гарнитура").replace
                                    ("выключатель", "выключатель/розетка").replace("Монтаж выключателя, Демонтаж выключателя", "Монтаж выключателя/розетки, Демонтаж выключателя/розетки").replace
                                    ("Демонтаж выключателя, Монтаж выключателя","Монтаж выключателя/розетки, Демонтаж выключателя/розетки");
                            cell2=row2.createCell(result);
                            cell2.setCellType(Cell.CELL_TYPE_STRING);
                            cell2.setCellStyle(style);
                            cell2.setCellValue(s);
                            int windth2=(int)(60*1.14388)*256;
                            sheet2.setColumnWidth(result,windth2);
                            result+=1;
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            System.out.print(cell.getNumericCellValue()+" ");
                            cell2=row2.createCell(result);
                            cell2.setCellType(Cell.CELL_TYPE_FORMULA);
                            cell2.setCellStyle(style);
                            cell2.setCellValue(cell.getNumericCellValue());
                            int windth3=(int)(10*1.14388)*256;
                            sheet2.setColumnWidth(result,windth3);
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
