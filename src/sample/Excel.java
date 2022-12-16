package sample;

import com.sun.media.sound.InvalidFormatException;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.*;

public class Excel {

    private String filename;
    private Integer Sheet;
    private Integer Row;
    private Integer Column;
    private XSSFWorkbook WorkbookExcel;
    private XSSFSheet SheetExcel;
    private XSSFRow RowExcel;
    private XSSFCell CellExcel;

    public void createExcel(){

        WorkbookExcel = new XSSFWorkbook();
        SheetExcel = WorkbookExcel.createSheet("1");    //пока только умеет работать с одним листом

    }

    public void createExcel(String filename, Integer Sheet){    //работа с уже созданной книгой

        this.filename = filename;
        this.Sheet = Sheet;
        WorkbookExcel = openBookDirectly(filename);
        SheetExcel = WorkbookExcel.getSheetAt(Sheet);

    }

    public XSSFCell getCell(Integer Row, Integer Column){

        this.Row = Row;
        this.Column = Column;

        RowExcel = SheetExcel.getRow(Row);
        if(RowExcel == null) { RowExcel = SheetExcel.createRow(Row); }

        CellExcel = RowExcel.getCell(Column);
        if(CellExcel == null) { CellExcel = RowExcel.createCell(Column); }

        return CellExcel;
    }

    public void setCell(Integer Row, Integer Column, String data){

        getCell(Row, Column);

        CellExcel.setCellValue(data);

    }

    public void setCell(Integer Row, Integer Column, Integer data){

        getCell(Row, Column);

        CellExcel.setCellValue(data);

    }

    //записывает картинку в одну ячейку
    //ИСТОЧНИК - https://coderoad.ru/33712621/Как-поместить-изображение-в-ячейку-excel-java
    public void setImg(int Row, int Column, String PathImage){

        try {

            //FileInputStream obtains input bytes from the image file
            InputStream inputStream = new FileInputStream(PathImage);
            //Get the contents of an InputStream as a byte[].
            byte[] bytes = IOUtils.toByteArray(inputStream);
            //Adds a picture to the workbook
            int pictureIdx = WorkbookExcel.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
            //close the input stream
            inputStream.close();
            //Returns an object that handles instantiating concrete classes
            CreationHelper helper = WorkbookExcel.getCreationHelper();
            //Creates the top-level drawing patriarch.
            Drawing drawing = SheetExcel.createDrawingPatriarch();

            //Create an anchor that is attached to the worksheet
            ClientAnchor anchor = helper.createClientAnchor();

            //create an anchor with upper left cell _and_ bottom right cell
            anchor.setCol1(Column); //Column B
            anchor.setRow1(Row); //Row 3
            anchor.setCol2(Column + 1); //Column C
            anchor.setRow2(Row + 1); //Row 4

            //Creates a picture
            Picture pict = drawing.createPicture(anchor, pictureIdx);

            //Reset the image to the original size
            //pict.resize(); //don't do that. Let the anchor resize the image!

            //Create the Cell B3
            Cell cell = SheetExcel.createRow(2).createCell(1);

            //set width to n character widths = count characters * 256
            //int widthUnits = 20*256;
            //sheet.setColumnWidth(1, widthUnits);

            //set height to n points in twips = n * 20
            //short heightUnits = 60*20;
            //cell.getRow().setHeight(heightUnits);

            //Write the Excel file
            //FileOutputStream fileOut = null;
            //fileOut = new FileOutputStream("myFile.xlsx");
            //WorkbookExcel.write(fileOut);
            //fileOut.close();

        } catch (IOException ioex) {
        }

    }

    //записывает картинку в диапозон ячеек. RowLength-высота картинки, ColumnLength-ширина картинки(измеряется ячейками)
    public void setImg(int Row, int Column, int RowLength, int ColumnLength, String PathImage){

        try {

            //FileInputStream obtains input bytes from the image file
            InputStream inputStream = new FileInputStream(PathImage);
            //Get the contents of an InputStream as a byte[].
            byte[] bytes = IOUtils.toByteArray(inputStream);
            //Adds a picture to the workbook
            int pictureIdx = WorkbookExcel.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
            //close the input stream
            inputStream.close();
            //Returns an object that handles instantiating concrete classes
            CreationHelper helper = WorkbookExcel.getCreationHelper();
            //Creates the top-level drawing patriarch.
            Drawing drawing = SheetExcel.createDrawingPatriarch();

            //Create an anchor that is attached to the worksheet
            ClientAnchor anchor = helper.createClientAnchor();

            //create an anchor with upper left cell _and_ bottom right cell
            anchor.setCol1(Column); //Column B
            anchor.setRow1(Row); //Row 3
            anchor.setCol2(Column + ColumnLength); //Column C
            anchor.setRow2(Row + RowLength); //Row 4

            //Creates a picture
            Picture pict = drawing.createPicture(anchor, pictureIdx);

            Cell cell = SheetExcel.createRow(2).createCell(1);

            //set width to n character widths = count characters * 256
            //int widthUnits = 20*256;
            //sheet.setColumnWidth(1, widthUnits);

            //set height to n points in twips = n * 20
            //short heightUnits = 60*20;
            //cell.getRow().setHeight(heightUnits);

            //Write the Excel file
            //FileOutputStream fileOut = null;
            //fileOut = new FileOutputStream("myFile.xlsx");
            //WorkbookExcel.write(fileOut);
            //fileOut.close();

        } catch (IOException ioex) {
        }

    }

    public void Build(String outFilename){

        createBookDirectly(WorkbookExcel, outFilename);

    }

    private static void createBookDirectly(XSSFWorkbook book, String NEW_FILE_NAME) {

        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(new File(NEW_FILE_NAME));
            book.write(fileOut);
            fileOut.flush();
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static XSSFWorkbook openBookDirectly(String filename) {

        XSSFWorkbook book = null;
        try {
            book = (XSSFWorkbook) WorkbookFactory.create(new File(filename));
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return book;
    }

}
