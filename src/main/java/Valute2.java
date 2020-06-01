
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.Properties;
import javax.mail.*;
import javax.mail.internet.*;
import javax.swing.JOptionPane;


public class Valute2 {                                           //создали   класс
    private static Document getPege() throws IOException {       //создали метод
        String url="https://yandex.ru/news/quotes/1.html ";      //указали url сайта для скачивания информации
     //пишим что хотим получить url , прописываем время ожидания.
        Document pege= (Document) Jsoup.parse ( new URL (url),5000);
        return pege;
    }

    public static void main(String[] args) throws IOException {

        Document page =getPege ();                                     //запрашиваем страницу
      //запрашиваем из этой страницы таблицу где содержатся данные с курсом валют.
        Element tableVal=page.select ("table").first ();
//создаем файл excel и страницу
        Workbook by = new XSSFWorkbook ();
        Sheet sheet = by.createSheet ( "Лист 1" );
// извлекаем из страницы данные с датой , данными курса, индексом изменения
        Elements val=tableVal.select ("td[class=quote__date]" );//даты
        Elements val1=tableVal.select ("td[class=quote__value]" );//значение
        Elements val2=tableVal.select ("td[class=quote__change]" );//индекс

// заносим данные о дате в массив
        String date =val.select ( "td[class=quote__date]" ).text ();
        String str= date;
        String[] words = str.split ("[ ]");
//заносим данные о курсе в массив
        String znak =val1.select ( "td[class=quote__value]" ).text ();
        String str1= znak;
        String[] words1 = str1.split("[[ ]*|[//.]]");
//заносим данные значений индекса в массив
        String index = val2.select ( "td[class=quote__change]" ).text ();
        String str2= index;
        String[] words2 = str2.split("[[ ]*|[//.]]");


  //создаем числовой массив  и внесем в него данные из массива words1
        double vil1[]=new double[10];
        for(int i=0;i<10;i++){
            vil1[i]= Double.parseDouble( words1[i].replace(",",".") );

        }
        //создаем числовой массив  и внесем в него данные из массива words1
        double vil2[]=new double[10];
        for(int i=0;i<10;i++){
            vil2[i]= Double.parseDouble( words2[i].replace(",",".") );

        }

//создаем столбцы в файле excel

        Row row0 = sheet.createRow ( 1 );
        Row row1 = sheet.createRow ( 2);
        Row row2 = sheet.createRow ( 3 );
        Row row3 = sheet.createRow ( 4 );
        Row row4 = sheet.createRow ( 5 );
        Row row5 = sheet.createRow ( 6 );
        Row row6 = sheet.createRow ( 7 );
        Row row7 = sheet.createRow ( 8 );
        Row row8 = sheet.createRow ( 9 );
        Row row9 = sheet.createRow ( 10 );
        Row row10 = sheet.createRow ( 11 );
//заносим название 1 столбца в 1 ячейку
        Cell cell1 = row0.createCell (0);
        cell1.setCellValue ("Дата");

  //заносим данные с датой из массива в таблицу
        Cell cell2= row1.createCell (0);
        cell2.setCellValue (words[0]);
        Cell cell3 = row2.createCell (0);
        cell3.setCellValue (words[1]);
        Cell cell4 = row3.createCell (0);
        cell4.setCellValue (words[2]);
        Cell cell5 = row4.createCell (0);
        cell5.setCellValue (words[3]);
        Cell cell6 = row5.createCell (0);
        cell6.setCellValue (words[4]);
        Cell cell7 = row6.createCell (0);
        cell7.setCellValue (words[5]);
        Cell cell8 = row7.createCell (0);
        cell8.setCellValue (words[6]);
        Cell cell9 = row8.createCell (0);
        cell9.setCellValue (words[7]);
        Cell cell10 = row9.createCell (0);
        cell10.setCellValue (words[8]);
        Cell cell11 = row10.createCell (0);
        cell11.setCellValue (words[9]);
//заносим название 2 столбца в 1 ячейку
        Cell cell12 = row0.createCell (1);
        cell12.setCellValue ("Курс доллара ");
//заносим данные с курсом валют из массива в таблицу
        Cell cell13= row1.createCell (1);
        cell13.setCellValue (vil1[0]);
        Cell cell14 = row2.createCell (1);
        cell14.setCellValue (vil1[1]);
        Cell cell15 = row3.createCell (1);
        cell15.setCellValue (vil1[2]);
        Cell cell16 = row4.createCell (1);
        cell16.setCellValue (vil1[3]);
        Cell cell17 = row5.createCell (1);
        cell17.setCellValue (vil1[4]);
        Cell cell18 = row6.createCell (1);
        cell18.setCellValue (vil1[5]);
        Cell cell19 = row7.createCell (1);
        cell19.setCellValue (vil1[6]);
        Cell cell20 = row8.createCell (1);
        cell20.setCellValue (vil1[7]);
        Cell cell21 = row9.createCell (1);
        cell21.setCellValue (vil1[8]);
        Cell cell22 = row10.createCell (1);
        cell22.setCellValue (vil1[9]);
//заносим название 3 столбца в 1 ячейку
        Cell cell23 = row0.createCell (2);
        cell23.setCellValue ("Изменение курса");
//заносим данные с индексом изменения  из массива в таблицу
        Cell cell24= row1.createCell (2);
        cell24.setCellValue (vil2[0]);
        Cell cell25 = row2.createCell (2);
        cell25.setCellValue (vil2[1]);
        Cell cell26 = row3.createCell (2);
        cell26.setCellValue (vil2[2]);
        Cell cell27 = row4.createCell (2);
        cell27.setCellValue (vil2[3]);
        Cell cell28 = row5.createCell (2);
        cell28.setCellValue (vil2[4]);
        Cell cell29 = row6.createCell (2);
        cell29.setCellValue (vil2[5]);
        Cell cell30 = row7.createCell (2);
        cell30.setCellValue (vil2[6]);
        Cell cell31 = row8.createCell (2);
        cell31.setCellValue (vil2[7]);
        Cell cell32 = row9.createCell (2);
        cell32.setCellValue (vil2[8]);
        Cell cell33 = row10.createCell (2);
        cell33.setCellValue (vil2[9]);
     //задаем размер ширины столбца для красоты восприятия
        sheet.setColumnWidth ( 1,4000 );//задаем размер ячеек
        sheet.setColumnWidth ( 2,4500 );//задаем размер ячеек
        sheet.addMergedRegion ( new CellRangeAddress (0,0,0,2  ) );//объеденяем ячейки

//указываем название страницы excel и место ее размешение
            FileOutputStream fos = new FileOutputStream ( "C:/Users/o-fed/Desktop/Artem1.xlsx" );
            by.write ( fos );
            fos.close ();
//отправка электронного письма об изменении данных курса валют.
            Properties p=new Properties (  );
            p.put ( "mail.smtp.host", "smtp.yandex.ru" );
            p.put ( "mail.smtp.socketFactory.port",465 );
            p.put ( "mail.smtp.socketFactory.class","javax.net.ssl.SSLSocketFactory");
            p.put ( "mail.smtp.auth","true" );
            p.put ( "mail.smtp.port",465 );

        Session s=Session.getDefaultInstance ( p,
           new javax.mail.Authenticator (){
            protected PasswordAuthentication getPasswordAuthentication(){
                return new PasswordAuthentication ( "fedolejka@yandex.ru","Binocl1843+" );
            }} );
        try{
//прописываем три варианта письма в зависимости от изменения курса валюты
            //если курс валют вырос
            if(vil1[0]>vil1[1]){
            Message mess= new MimeMessage ( s );
            mess.setFrom(new InternetAddress ("fedolejka@yandex.ru"));
            mess.setRecipients ( Message.RecipientType.TO,InternetAddress.parse ( "o-fedoreev@mail.ru" ) );
            mess.setSubject ( " Данные курса валют" );
            mess.setText ( "Курс вырос" );
            Transport.send ( mess );}
            //если курс валют не изменился
            else if(vil1[0]==vil1[1]){
                Message mess= new MimeMessage ( s );
                mess.setFrom(new InternetAddress ("fedolejka@yandex.ru"));
                mess.setRecipients ( Message.RecipientType.TO,InternetAddress.parse ( "o-fedoreev@mail.ru" ) );
                mess.setSubject ( " Данные курса валют" );
                mess.setText ( "Курс не изменился" );
                Transport.send ( mess );}
             //если курс валют упал
            else {

                Message mess= new MimeMessage ( s );
                mess.setFrom(new InternetAddress ("fedolejka@yandex.ru"));
                mess.setRecipients ( Message.RecipientType.TO,InternetAddress.parse ( "o-fedoreev@mail.ru" ) );
                mess.setSubject ( "Данные курса валют" );
                mess.setText ( "Курс упал");
                Transport.send ( mess );}
        }
        catch (Exception ex){
            JOptionPane.showMessageDialog ( null,"Что то пошло не так"+ex );
        } }}









