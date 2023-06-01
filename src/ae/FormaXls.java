/*
 * Copyright (c) 2017. Aleksey Eremin
 * 10.02.17 14:41
 * 04.09.19
 *
 * Формирование листа Excel по данным из БД
 *
 */

package ae;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

class FormaXls {
    private Database f_db;
    
    /**
     * Конструктор
     * @param db    база данных, с таблицей ALERTS
     */
    FormaXls(Database db) {
        this.f_db = db;
    }

    /**
     * Изготовить лист по отчета по Рейтингу
     * @param year      год
     * @param month     месяц, если 0, значит отчет за год
     * @param day       день, если 0, значит отчет за месяц
     * @return          имя сформированного файла (в каталоге excelDir)
     */
    String makeList(int year, int month, int day)
    {
        final int Data_base_row = 2;       // базовая строка, для вставки данных
        final int Date_base_col = 1;       // базовая колонка для вставки данных

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
            // заполним лист данными за требуемую дату

            String strDat1 = String.format("%04d-%02d-%02d", year, month, day); // дата рейтинга
            //
            String sql;
            ArrayList<String[]> arrlst;
            // SELECT / pn, DATE_FORMAT(Rating.dat,'%d.%m.%Y'),
            sql = "SELECT " +
                "1," +                                          // порядковый номер строки
                "Regions.nam, " +                            // название региона
                "Rating.op_name, " +                            // оператор
                "Rating.id_region, " +                          // номер региона
                "Rating.inn, " +                                // ИНН
                "CONCAT(ROUND(100*not_block/total,3),''), " +   // процент (не ставим знак %)
                "total, " +                                     // всего в реестре
                "not_block, " +                                 // не заблокировано
                "GROUP_CONCAT(note SEPARATOR ' | ') " +         // вышестоящие
                "FROM Rating LEFT JOIN " +
                "(Opers LEFT JOIN opnotes ON (Opers.op_id=opnotes.op_id AND opnotes.tip='Uplink')) " +
                "ON Rating.inn = Opers.op_inn " +
                "LEFT JOIN Regions ON Rating.id_region=Regions.id " +
                "WHERE Rating.dat='" + strDat1 + "' " +
                "GROUP BY pn;";
            arrlst = f_db.DlookupArray(sql);
            int cnt = 0; // кол-во записанных строк
            for(String[] rst: arrlst) {
                Row row = wks.getRow(Data_base_row + cnt);
                if(row == null) {
                    row = wks.createRow(Data_base_row + cnt);
                }
                // преобразовать формат даты на Java
//                DateTimeFormatter format = DateTimeFormatter.ofPattern("yyyy-MM-dd");
//                LocalDateTime dt = LocalDateTime.parse(sDat, format);
//                String sDatO = dt.format(DateTimeFormatter.ofPattern("dd.MM.YYYY"));
                cnt++;
                rst[0] = Integer.toString(cnt); // порядковый номер строки
                setRowVals(row, rst); // записать строку
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
     * Записать значения в строку Excel из массива строк ответа БД
     * и преобразованием некоторых позиций в целое
     * @param row   строка Excel, куда делается запись
     * @param rst   массив строк для записи
     */
    private void setRowVals(Row row, String[] rst)
    {
        final String intIndex = R.intIndex;     // "(0)(3)(6)(7)"; // список колонок с целыми числами
        final String dblIndex = R.dblIndex;     //"(5)"; // список колонок с действительнымии числами
        //
        for(int i = 0; i < rst.length; i++) {
            String r = rst[i];
            if(intIndex.contains("("+i+")")) {
                // числовая колонка
                try {
                    int v = Integer.parseInt(r); // числовое представление
                    setCellVal(row, i, v);
                } catch (Exception e) {
                    System.err.println("Ошибка преобразования числа: " + r + " - " + e.getMessage());
                }
            }else if(dblIndex.contains("("+i+")")) {
                // действительная колонка
                try {
                    double v = Double.parseDouble(r); // числовое представление
                    setCellVal(row, i, v);
                } catch (Exception e) {
                    System.err.println("Ошибка преобразования числа: " + r + " - " + e.getMessage());
                }
            } else {
                setCellVal(row, i, r);
            }
        }
    }

    /**
     * Установить действительное числовое значение ячейки в заданной строке таблицы
     * @param row   строка
     * @param col   номер колонки
     * @param val   устанавливаемое значения (double)
     * @return      1 - значение установлено, 0 - не установлено
     */
    private boolean setCellVal(Row row, int col, double val)
    {
        try {
            getCell(row, col).setCellValue(val);  // Access the cell
        } catch (Exception e) {
            System.err.println("ошибка здания значения клетке " + col + " value: " + val);
            return false;
        }
        return true;
    }

    /**
     * Установить числовое значение ячейки в заданной строке таблицы
     * @param row   строка
     * @param col   номер колонки
     * @param val   устанавливаемое значения (long)
     * @return      1 - значение установлено, 0 - не установлено
     */
    private boolean setCellVal(Row row, int col, int val)
    {
        try {
            getCell(row, col).setCellValue(val);  // Access the cell
        } catch (Exception e) {
            System.err.println("ошибка здания значения клетке " + col + " value: " + val);
            return false;
        }
        return true;
    }

    /**
     * Установить строковое значение ячейки в заданной строке таблицы
     * @param row   строка
     * @param col   номер колонки
     * @param val   устанавливаемое значения (String)
     * @return      1 - значение установлено, 0 - не установлено
     */
    private boolean setCellVal(Row row, int col, String val)
    {
        try {
            getCell(row, col).setCellValue(val);  // Access the cell
        } catch (Exception e) {
            System.err.println("ошибка здания значения клетке " + col + " value: " + val);
            return false;
        }
        return true;
    }

    /**
     * Получить ячейки в строке в заданной колонке
     * @param row   строка
     * @param col   индекс колонки
     * @return  ячейка (клетка)
     */
    private Cell getCell(Row row, int col)
    {
        Cell c = row.getCell(col);  // Access the cell
        if (c == null) {
            c = row.createCell(col); // создадим ячейку
        }
        return c;
    }
    
} // end of class
