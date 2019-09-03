/*
 * Copyright (c) 2017. Aleksey Eremin
 * 10.02.17 14:41
 */

package ae;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;

/**
 * Created by ae on 10.02.2017.
 * Формирование Формы 10 на основе данных таблицы ALERTS
 */
public class FormaXls {
    private Database f_db;
    
    /**
     * Конструктор
     * @param db    база данных, с таблицей ALERTS
     */
    public FormaXls(Database db) {
        this.f_db = db;
    }

    /**
     * Изготовить лист по отчета по Рейтингу
     * @param year      год
     * @param month     месяц, если 0, значит отчет за год
     * @param day       день, если 0, значит отчет за месяц
     * @return          имя сформированного файла (в каталоге excelDir)
     */
    public String makeList(int year, int month, int day)
    {
        final int Data_base_row = 2;       // базовая строка, для вставки данных
        final int Date_base_col = 1;       // базовая колонка для вставки данных

        int Cnt = 0; // кол-во записанных строк
        // Сохранить в каталоге отчетов Excel с Формой 10 за указанное число
        R r = new R();  // ресурсный файл
        String resname = "res/" + R.fileNameExcel;
        String s = String.format("%04d%02d%02d_",year,month,day);
        String fileName = R.workDir + R.sep + s + R.fileNameExcel;
        if(!r.writeRes2File(resname, fileName)) {
            System.out.println("?ERROR-Can't write file: " + fileName);
            return null;
        }
        //
        try {
            FileInputStream inp = new FileInputStream(fileName);
            // получим рабочую книгу Excel
            //Workbook wb = new XSSFWorkbook(inp); // прочитать файл с Excel 2010
            HSSFWorkbook wb = new HSSFWorkbook(inp); // прочитать файл с Excel 2003
            inp.close();
            // Read more: http://www.techartifact.com/blogs/2013/10/update-or-edit-existing-excel-files-in-java-using-apache-poi.html#ixzz4Y23Vf1eR
            // получим первый лист
            HSSFSheet wks = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
            // заполним лист данными
            // если указаны требуемая дата, выберем ее
            // если указан год, выбираем все за год,
            // если указан еще и месяц выберем год и месяц,
            // если указан еще и день, то выберем год, месяц и день.
            // Если указан час (hour >= 0), то считаем дату полностью заданной и час тоже

            String strDat1 = String.format("%04d-%02d-%02d", year, month, day); // дата рейтинга

//            if(hour < 0) {
//                // выбираем события за указаный год, месяц, день
//                if(year != 0) {
//                    String st = "%Y";
//                    String sf = "%04d";
//                    strDat1 =String.format("%04d-01-01", year);
//                    if (month != 0) {
//                        st = st + "-%m";
//                        sf = sf + "-%02d";
//                        strDat1 =String.format("%04d-%02d-01", year, month);
//                        if (day != 0) {
//                            st = st + "-%d";
//                            sf = sf + "-%02d";
//                            strDat1 =String.format("%04d-%02d-%02d", year, month, day);
//                        }
//                    }
//                    String ssdat = String.format(sf, year, month, day);
//                    strBetw = "(strftime('" + st + "',Dat)='" + ssdat + "')";
//                    strDat1 = "datetime('" + strDat1 + "')";
//                }
//            } else {
//                // события за сутки до указанного часа
//                String sh = String.format("%04d-%02d-%02d %02d:00:00", year, month, day); // нужный час
//                strBetw = "(Dat BETWEEN datetime('" + sh +"', '-24 hours') AND datetime('" + sh + "'))";
//                strDat1 = "datetime('" + sh + "', '-24 hours')";
//            }
            //
            String sql;
            ArrayList<String[]> arrlst;
            // SELECT
            sql = "SELECT pn, DATE_FORMAT(Rating.dat,'%d.%m.%Y'), Rating.id_region, " +
                "Rating.inn, Rating.op_name, " +
                "total, not_block, CONCAT(ROUND(100*not_block/total,2),'%') as prct, " +
                "GROUP_CONCAT(note SEPARATOR ' | ') as uplink " +
                "FROM Rating LEFT JOIN " +
                "(Opers LEFT JOIN opnotes ON (Opers.op_id=opnotes.op_id AND opnotes.tip='Uplink')) " +
                "ON Rating.inn = Opers.op_inn " +
                "WHERE Rating.dat='" + strDat1 + "' " +
                "GROUP BY pn,Rating.dat, Rating.id_region, Rating.inn, Rating.op_name,total,not_block;";
            arrlst = f_db.DlookupArray(sql);
            Cnt = 0;
            for(String[] rst: arrlst) {
//                String sDat = rst[1];       // взять дату
//                String sReg = rst[2];       // регион
//                String sInn = rst[3];       // ИНН
//                String sOpr = rst[4];       // оператор
//                String sTot = rst[5];       // всего записей реестра
//                String sNbl = rst[6];       // не заблокировано
//                String sPrc = rst[7]+"%";   // процент незаблокированных
//                String sUpe = rst[8];       // вышестоящие операторы
                // базовая строка для вставки данных для данного числа месяца
                Row row = wks.getRow(Data_base_row + Cnt);
                if(row == null) {
                    row = wks.createRow(Data_base_row + Cnt);
                }
                // преобразовать формат даты на Java
//                DateTimeFormatter format = DateTimeFormatter.ofPattern("yyyy-MM-dd");
//                LocalDateTime dt = LocalDateTime.parse(sDat, format);
//                String sDatO = dt.format(DateTimeFormatter.ofPattern("dd.MM.YYYY"));
                setRowVals(row, rst); // записать строку
                Cnt++;
            }
            //
            // установить дату на листе
            strDat1 =String.format("%02d.%02d.%04d", day, month, year);
            // ячейка даты
            Row row = wks.getRow(0);
            setCellVal(row, 2, strDat1);
            // После заполнения ячеек формулы не пересчитываются, поэтому выполним принудительно
            // перерасчет всех формул на листе
            // http://poi.apache.org/spreadsheet/eval.html#Re-calculating+all+formulas+in+a+Workbook
            // в данной задаче в листе Excel нет формул, поэтому этот код ниже закоментирован
            //// FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            //// for (Sheet sheet : wb) { for (Row row : sheet) {  for (Cell c : row) { if (c.getCellType() == Cell.CELL_TYPE_FORMULA) { evaluator.evaluateFormulaCell(c); }  }  } }
            //
            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream(fileName);
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
        //
        return fileName;
    }

    /**
     * Записать строку в Excel c 1 gj 8 ячейки массива
     * @param row   строка Excel, куда делается запись
     * @param rst   массив строк для записи
     */
    private void setRowVals(Row row, String[] rst)
    {
        for(int i = 0; i < 8; i++) {
            setCellVal(row, i, rst[1+i]);
        }
    }

//    /**
//     * Установить числовое значение ячейки в заданной строке таблицы
//     * @param row   строка
//     * @param col   номер колонки
//     * @param val   устанавливаемое значения (int)
//     * @return      1 - значение установлено, 0 - не установлено
//     */
//    private int setCellVal(Row row, int col, int val)
//    {
//        return setCellVal(row, col, (long)val);
//    }
//
//    /**
//     * Установить числовое значение ячейки в заданной строке таблицы
//     * @param row   строка
//     * @param col   номер колонки
//     * @param val   устанавливаемое значения (long)
//     * @return      1 - значение установлено, 0 - не установлено
//     */
//    private int setCellVal(Row row, int col, Long val)
//    {
//        Cell c = row.getCell(col);  // Access the cell
//        if (c == null) {
//            c = row.createCell(col); // создадим ячейку
//        }
//        if(c != null) {
//            c.setCellValue(val);
//            return 1;
//        }
//        return 0;
//    }
    
    /**
     * Установить строковое значение ячейки в заданной строке таблицы
     * @param row   строка
     * @param col   номер колонки
     * @param val   устанавливаемое значения (String)
     * @return      1 - значение установлено, 0 - не установлено
     */
    private int setCellVal(Row row, int col, String val)
    {
        Cell c = row.getCell(col);  // Access the cell
        if (c == null) {
            c = row.createCell(col); // создадим ячейку
        }
        if(c != null) {
            c.setCellValue(val);
            return 1;
        }
        return 0;
    }
    
    
} // end of class

