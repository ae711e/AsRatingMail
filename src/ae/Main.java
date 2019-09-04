/*
 * Copyright (c) 2019. Aleksey Eremin
 * 03.09.19 11:39
 */
/*
 * Отправка по почте рейтинга операторов за вчерашнее число, формируемого на основе данных MySql
 *
 */
package ae;

import java.io.File;
import java.time.LocalDateTime;

public class Main {

    public static void main(String[] args) {
	      // write your code here
        System.out.println("Данные рейтинга");

        if(!prepareWork()) {
            System.out.println("Error prepare work");
            return;
        }
        // вчерашняя дата
        final LocalDateTime dt = LocalDateTime.now().minusHours(24);
        int d = dt.getDayOfMonth();
        int m = dt.getMonthValue();
        int y = dt.getYear();
        String sdat = String.format("%02d.%02d.%04d",d,m,y);
        //
        System.out.println("Рейтинг на "+ sdat);

        // откроем БД
        Database db = new DatabaseMysql(R.dbHost,R.dbBase,R.dbUser,R.dbPass);
        //
        // создадим объект для формирования отчета Excel
        FormaXls f = new FormaXls(db);

        String outFile = f.makeList(y,m,d);
        //
        MailSend mc = new MailSend();
        String  otv;
        otv = mc.mailSend(R.SmtpMailTo,R.MailSubject + " " + sdat, R.MailMessage, outFile);
        if(otv != null) {
            System.out.println("Почта отправлена: " + R.SmtpMailTo);
        }
    }

    /**
     * Подготовить данные и файлы для работы
     * @return true - подготовлено, false - не готово
     */
    private static boolean prepareWork()
    {
        R r = new R();   // для загрузки ресурсных значений
        r.loadDefault(); // значения по умолчанию
        //
        System.out.println("work  dir: " + R.workDir);
        //
        // проверим наличие каталогов
        String[] dstr = new String[]{R.workDir};
        for(int i=0; i < dstr.length; i++) {
            String s = dstr[i];
            File f = new File(s);
            if(!f.exists()) {
                System.out.println("?ERROR-Not found dir: " + s);
                return false;
            }
        }
        //
        return true;
    }

}
