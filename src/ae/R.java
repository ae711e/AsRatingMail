/*
 * Copyright (c) 2017. Aleksey Eremin
 * 28.01.17 21:52
 * 04.09.19
 */

package ae;

import java.io.*;
import java.util.Properties;

/**
 * Created by ae on 28.01.2017.
 * Отправка письма с таблицей рейтинга, формируемого на основе данных MySql
 * Ресурсный класс
 *
 * Modify:
 * 04.09.19  изменил форму таблицы
 *
 */
public class R {
    private final static String Ver = "Ver. 1.03"; // номер версии
    
    final static String sep = System.getProperty("file.separator"); // разделитель имени каталогов

    static String workDir =   null; // рабочий каталог. если null, возьмет системный временный каталог

    final static String fileNameExcel = "rating.xls";  // имя файла Excel

    static String   dbHost = _r.dbHost ;// сервер MySql
    static String   dbBase = _r.dbBase ;// база данных MySql
    static String   dbUser = _r.dbUser ;// пользоватекль MySql
    static String   dbPass = _r.dbPass ;// пароль MySql
    // почтовые дела
    // адрес получателя почты (можно несколько с разделением по ;)
    static String SmtpMailTo  = _r.smtpmailcc;  // адрес(а) получателя почты
    static String SmtpMailCC  = _r.smtpmailcc;  // адрес получателя копии почты
    
    static String SmtpServer     = _r.smtpserver;            // адрес почтового сервера
    static int    SmtpServerPort = _r.smtpserverport;                      // порт почтового сервера
    static String SmtpSender     = _r.smtpsender; // адрес отправителя почты
    static String SmtpServerUser = _r.smtpserveruser;                    // имя пользователя почтового сервера
    static String SmtpServerPwd  = _r.smtpserverpwd;                     // пароль пользователя почтового сервера

    // тема и тело письма "Отчет о контроле доступности"
    static String MailSubject = "Рейтинг";
    static String MailMessage = "Рейтинг операторов. Приложение 1. Файл Excel";
    
    /**
     * загрузка значений параметров по-умолчанию из файла res/default.properties
     * Порядок определения каталогов:
     * ОБЯЗАТЕЛЬНО ДОЛЖЕН БЫТЬ workDir
     * если не задан screenDir (null и закоментирован в properties) берем как workDir
     * если не задан excelDIr берем как screenDir
     * То есть, можно задать только workDir и все будет валится туда
     */
    void loadDefault()
    {
        // http://stackoverflow.com/questions/2815404/load-properties-file-in-jar
        // Отобразим версию
        System.out.println(Ver);
        Properties props = new Properties();
        try {
            props.load(R.class.getResourceAsStream("res/default.properties"));
            // прочитаем параметры из конфигурационного файла default.properties
            workDir = r2s(props, "workDir", workDir);
            if(workDir == null) {
                workDir = System.getProperty("java.io.tmpdir", ".");
            }
            // почта
            SmtpServer = r2s(props, "SmtpServer", SmtpServer);
            SmtpMailTo = r2s(props, "SmtpMailTo", SmtpMailTo);
            SmtpMailCC = r2s(props, "SmtpMailCC", SmtpMailCC);
            // база данных
            dbHost = r2s(props, "dbHost", dbHost);
            dbBase = r2s(props, "dbBase", dbBase);
            dbUser = r2s(props, "dbUser", dbUser);
            dbPass = r2s(props, "dbPass", dbPass);
            //
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Выдать строковое значение из файла свойств, либо, если там
     * нет такого свойства, вернуть значение по-умолчанию
     * @param p                     свойства
     * @param NameProp              имя свойства
     * @param strResourceDefault    значение по-умолчанию
     * @return  значение свойства, а если его нет, то значение по-умолчанию
     */
    private String r2s(Properties p, String NameProp, String strResourceDefault)
    {
        String str = p.getProperty(NameProp);
        if(str == null) {
            str = strResourceDefault;
        }
        return str;
    }

    /**
     * Выдать числовое (long) значение из файла свойств, либо, если там
     * нет такого свойства, вернуть значение по-умолчанию
     * @param p                     свойства
     * @param NameProp              имя свойства
     * @param lngResourceDefault    значение по-умолчанию
     * @return  значение свойства, а если его нет, то значение по-умолчанию
     */
    private long r2s(Properties p, String NameProp, long lngResourceDefault)
    {
        String str = p.getProperty(NameProp);
        if(str == null) {
            str = String.valueOf(lngResourceDefault);
        }
        return Long.parseLong(str);
    }

    /**
     * Выдать числовое (int) значение из файла свойств, либо, если там
     * нет такого свойства, вернуть значение по-умолчанию
     * @param p                     свойства
     * @param NameProp              имя свойства
     * @param intResourceDefault    значение по-умолчанию
     * @return  значение свойства, а если его нет, то значение по-умолчанию
     */
    private int r2s(Properties p, String NameProp, int intResourceDefault)
    {
        String str = p.getProperty(NameProp);
        if(str == null) {
            str = String.valueOf(intResourceDefault);
        }
        return Integer.parseInt(str);
    }

    /**
     * прочитать ресурсный файл
     * by novel  http://skipy-ru.livejournal.com/5343.html
     * https://docs.oracle.com/javase/tutorial/deployment/webstart/retrievingResources.html
     * @param nameRes - имя ресурсного файла
     * @return - содержимое ресурсного файла
     */
    public String readRes(String nameRes)
    {
        String str = null;
        ByteArrayOutputStream buf = readResB(nameRes);
        if(buf != null) {
            str = buf.toString();
        }
        return str;
    }

    /**
     * Поместить ресурс в байтовый массив
     * @param nameRes - название ресурса (относительно каталога пакета)
     * @return - байтовый массив
     */
    private ByteArrayOutputStream readResB(String nameRes)
    {
        String str = null;
        try {
            // Get current classloader
            InputStream is = getClass().getResourceAsStream(nameRes);
            if(is == null) {
                System.out.println("Not found resource: " + nameRes);
                return null;
            }
            // https://habrahabr.ru/company/luxoft/blog/278233/ п.8
            BufferedInputStream bin = new BufferedInputStream(is);
            ByteArrayOutputStream bout = new ByteArrayOutputStream();
            int len;
            byte[] buf = new byte[512];
            while((len=bin.read(buf)) != -1) {
                bout.write(buf,0,len);
            }
            return bout;
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        return null;
    }

    /**
     * Записать в файл текст из строкт
     * @param strTxt - строка текста
     * @param fileName - имя файла
     * @return      true - записано, false - ошибка
     */
    public boolean writeStr2File(String strTxt, String fileName)
    {
        File f = new File(fileName);
        try {
            // сформируем командный файл BAT
            PrintWriter out = new PrintWriter(f);
            out.write(strTxt);
            out.close();
        } catch(IOException ex) {
            ex.printStackTrace();
            return false;
        }
        return true;
    }

    /**
     *  Записать в файл ресурсный файл
     * @param nameRes   имя ресурса (от корня src)
     * @param fileName  имя файла, куда записывается ресурс
     * @return  true - запись выполнена, false - ошибка
     */
    boolean writeRes2File(String nameRes, String fileName)
    {
        boolean b = false;
        ByteArrayOutputStream buf = readResB(nameRes);
        if(buf != null) {
            try {
                FileOutputStream fout = new FileOutputStream(fileName);
                buf.writeTo(fout);
                fout.close();
                b = true;
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return b;
    }
    
    /**
     * Загружает текстовый ресурс в заданной кодировке
     * @param name      имя ресурса
     * @param code_page кодировка, например "Cp1251"
     * @return          строка ресурса
     */
    public String getText(String name, String code_page)
    {
        StringBuilder sb = new StringBuilder();
        try {
            InputStream is = this.getClass().getResourceAsStream(name);  // Имя ресурса
            BufferedReader br = new BufferedReader(new InputStreamReader(is, code_page));
            String line;
            while ((line = br.readLine()) !=null) {
                sb.append(line);  sb.append("\n");
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        return sb.toString();
    }
    
    /**
     * Пауза выполнения программы (потока)
     * @param msec - задержка, мсек
     */
    public static void Sleep(long msec)
    {
        try {
            Thread.sleep(msec);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

} // end of class
