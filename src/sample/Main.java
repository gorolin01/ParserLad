package sample;

import java.awt.*;
import java.io.*;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
/*import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;*/

public class Main {

    private static String SiteName = "Lad";
    private static String mainUrl = "http://ptplad.ru";
    private static int fromPage = 1;
    private static int toPage = 1;
    private  static int NUMBER_OF_COLUMNS = 5;

    public static void main(String[] args) {

       /* System.setProperty("webdriver.chrome.driver", "selenium\\chromedriver.exe");
        WebDriver webDriver = new ChromeDriver();
        webDriver.get("https://yug-instrument.ru/catalog/elektroinstrumenty/pily/pily_montazhnye_otreznye/8869528/");
        webDriver.getPageSource();*/

        //ССЫЛКИ СТРАНИЦ С ТОВАРОМ НУЖНО ЗАКИНУТЬ В ФАЙЛ pageForParsing.txt(ССЫЛКИ ПОИСКА)

        try {
            parser(mainUrl, fromPage, toPage);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static Document getDoc(String url) {

        Excel excel = new Excel();
        excel.createExcel();

        System.out.println("Connect to page...");
        Connection connect = Jsoup.connect(url)
                .userAgent("Mozilla");
        boolean connected=false;
        Document doc=null;
        while(!connected){
            try{
                doc = connect.get();
                connected=true;
            }catch(Exception ex){

            }finally{
                System.out.println("connected: "+connected);
                if(!connected){
                    try{
                        Thread.sleep(1000);
                    }catch(Exception ex){

                    }
                }
            }
        }
        System.out.println("Ok!");

        return doc;

    }

    public static ArrayList<String> getArrayStrOnFile(String pathname) {
        ArrayList<String> Data = new ArrayList<>();
        try {
            File file = new File(pathname);
            FileReader fr = new FileReader(file);
            BufferedReader reader = new BufferedReader(fr);

            String line = reader.readLine();
            while (line != null) {
                Data.add(line);
                line = reader.readLine();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return Data;
    }

    public static String [] corrector(String title, String attribute) {
        String metric;
        String [] res = new String[2];

        if(title.contains(",")) {
            metric = title.substring(title.indexOf(",") + 2);
            title = title.substring(0, title.indexOf(","));
            attribute = attribute + " " + metric;
        }
        title = title + ":";

        res[0] = title;
        res[1] = attribute;

        return res;
    }

    private static void Download(String URL, String Name, String URLSave) throws Exception {

        try{
            String fileName = Name;
            String website = URL;

            System.out.println("Downloading File From: " + website);

            java.net.URL url = new URL(website);
            InputStream inputStream = url.openStream();
            OutputStream outputStream = new FileOutputStream(URLSave + "/" + fileName);
            byte[] buffer = new byte[2048];

            int length = 0;

            while ((length = inputStream.read(buffer)) != -1) {
                System.out.println("Buffer Read of length: " + length);
                outputStream.write(buffer, 0, length);
            }

            inputStream.close();
            outputStream.close();

        } catch(Exception e) {
            System.out.println("Exception: " + e.getMessage());
        }

    }

    //запись в txt файл
    private static void writeOnTxt(String data, String path, int noOfLines) {
        File file = new File(path);
        FileWriter fr = null;
        BufferedWriter br = null;
        String dataWithNewLine = data + System.getProperty("line.separator");
        try{
            fr = new FileWriter(file);
            br = new BufferedWriter(fr);
            for(int i = noOfLines; i>0; i--){
                br.write(dataWithNewLine);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
            try {
                br.close();
                fr.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static void writeTXT(String data, String src)
    {
        Path path = Paths.get(src);

        String OldFileData = "";
        File file = new File(src);
        if(file.exists()){
            try {
                BufferedReader reader = new BufferedReader(new FileReader(file));
                // считаем сначала первую строку
                String line = reader.readLine();
                while (line != null) {
                    OldFileData += line + "\n";
                    // считываем остальные строки в цикле
                    line = reader.readLine();
                }
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
            //OldFileData += "\n";
        }

        try (BufferedWriter bw = Files.newBufferedWriter(path, StandardCharsets.UTF_8))
        {
            bw.write(OldFileData);
            bw.write(data);
            System.out.println("Successfully written data to the file");
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void parser(String mainUrl, int fromPage, int toPage) throws IOException, NullPointerException {

        Elements select;

        Excel excel = new Excel();
        excel.createExcel();
        int Row = 0;
        int Column = 0;

        //System.setProperty("webdriver.chrome.driver", "selenium\\chromedriver.exe");

        new File("parsing_" + SiteName).mkdir();
        ArrayList<String> URLPage = getArrayStrOnFile("pageForParsing.txt"); //тут указываем ссылки на страници для парсинга

        for(int w = 0; w < URLPage.size(); w++) {

            Document doc = getDoc(URLPage.get(w));
            ArrayList<String> resList;

            //получили ссылки на страници товаров из меню товаров
            /*ArrayList<String> ListKartochkiInPage = new ArrayList<>();
            for(int KartochkiInPage = 0; KartochkiInPage < doc.select(".item-title").size(); KartochkiInPage++){
                System.out.println(mainUrl + doc.select(".item-title").get(KartochkiInPage).select("a").attr("href"));
                ListKartochkiInPage.add(mainUrl + doc.select(".item-title").get(KartochkiInPage).select("a").attr("href"));
            }*/

            for(int nomerTovara = 0; nomerTovara < doc.select(".col4").size(); nomerTovara++){

                Row = (Column / NUMBER_OF_COLUMNS) * NUMBER_OF_COLUMNS;

                //загрузка картинок
                String s = doc.select(".img_wrap").get(nomerTovara).select("img").attr("src");
                System.out.println(s);
                if(s.contains("/")){
                    try {
                        Download(mainUrl + s, doc.select(".code").select(".left").get(nomerTovara).text().replace(":", "") , "parsing_" + SiteName);
                    } catch (Exception exception) {
                        exception.printStackTrace();
                    }
                }
                excel.setImg(Row, Column % NUMBER_OF_COLUMNS,"parsing_" + SiteName + "/" + doc.select(".code").select(".left").get(nomerTovara).text().replace(":", ""));
                Row++;

                //название
                excel.setCell(Row, Column % NUMBER_OF_COLUMNS, doc.select(".product_info").get(nomerTovara).select(".ttl").text());
                Row++;

                //артикул
                excel.setCell(Row, Column % NUMBER_OF_COLUMNS, doc.select(".code").select(".left").get(nomerTovara).text().substring(11));
                Row++;

                //Штрих-код
                excel.setCell(Row, Column % NUMBER_OF_COLUMNS, "ШК");
                Row++;

                Column++;
            }

        }

        System.out.println(URLPage.size());
        excel.Build("parsing_" + SiteName + "/" + "Описание" + ".xlsx");
    }

}
