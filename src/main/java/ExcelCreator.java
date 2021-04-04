
import me.tongfei.progressbar.ProgressBar;
import me.tongfei.progressbar.ProgressBarStyle;
import org.apache.poi.ss.formula.functions.Match;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.util.ArrayList;
import java.util.Locale;


/**
 * Created on 10.03.2021
 *
 * @author Mykola Horkov
 * mykola.horkov@gmail.com
 */

public class ExcelCreator {

    public static void instanceToExcelFromTemplate(ArrayList<Manifest> manifests, String pathToSave) {
        if (manifests.size() == 0){
            System.out.println("No manifests were found...");
            return;
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            XSSFSheet sheet = workbook.createSheet("Manifests");
            sheet.createRow(0);
            sheet.createRow(1);

            sheet.getRow(1).createCell(0);
            sheet.getRow(1).getCell(0).setCellValue("Pick-up Date");

            sheet.getRow(1).createCell(1);
            sheet.getRow(1).getCell(1).setCellValue("Delivery Date");

            sheet.getRow(1).createCell(2);
            sheet.getRow(1).getCell(2).setCellValue("Plant");

            sheet.getRow(1).createCell(3);
            sheet.getRow(1).getCell(3).setCellValue("Invert. Date");

            sheet.getRow(1).createCell(4);
            sheet.getRow(1).getCell(4).setCellValue("Manifest");

            sheet.getRow(1).createCell(5);
            sheet.getRow(1).getCell(5).setCellValue("Supplier");

            sheet.getRow(1).createCell(6);
            sheet.getRow(1).getCell(6).setCellValue("Pallet QTY");

            sheet.getRow(1).createCell(7);
            sheet.getRow(1).getCell(7).setCellValue("m3");

            sheet.getRow(1).createCell(8);
            sheet.getRow(1).getCell(8).setCellValue("Weight");

            int rowIdx = 2;


            System.out.println(manifests.size() + " manifests were found...");
            ProgressBar pb = new ProgressBar("Progress pdf", manifests.size(), ProgressBarStyle.ASCII).start();

            for (int i = 0; i < manifests.size(); i++, rowIdx++) {
                Row row = sheet.createRow(rowIdx);
                fillRowWithData(manifests.get(i), row);
                pb.step();
                pb.setExtraMessage("Parsing...");
            }
            pb.stop();

            System.out.println("Writing of Excel file");
            workbook.write(out);
            try (FileOutputStream outputStream = new FileOutputStream(pathToSave)) {
                workbook.write(outputStream);
                System.out.println("DONE!");
            } catch (IOException en) {
                System.out.println(en.getMessage());
            }
        } catch (IOException e) {
            System.out.println(String.format("Error occurred while creating file {}", e.getMessage()));
            e.printStackTrace();
        }
    }

    public static void fillRowWithData(Manifest manifest, Row row) {

        Cell pickUpDate = row.createCell(0);
        pickUpDate.setCellValue(parseDate(manifest.getDateCollect()));

        Cell deliveryDate = row.createCell(1);
        deliveryDate.setCellValue(parseDate(manifest.getDateDeliver()));

        Cell plantCell = row.createCell(2);
        plantCell.setCellValue(manifest.getPlant());

        Cell invertDate = row.createCell(3);
        invertDate.setCellValue(manifest.getInvertedDate());

        Cell manifestCell = row.createCell(4);
        manifestCell.setCellValue(manifest.getManifest());

        Cell supplierCell = row.createCell(5);
        supplierCell.setCellValue(manifest.getSupplier());

        Cell palletQtyCell = row.createCell(6);
        palletQtyCell.setCellValue(manifest.getPalletQty());

        Cell m3Cell = row.createCell(7);
        m3Cell.setCellValue(manifest.getQbm());

        Cell weightCell = row.createCell(8);
        weightCell.setCellValue(manifest.getWeight());
    }

    private static String parseDate(String dateS) {
        DateTimeFormatterBuilder builder = new DateTimeFormatterBuilder()
                .parseCaseInsensitive().parseLenient()
                .appendPattern("dd-MMM-yyyy");

        DateTimeFormatter parser = builder.toFormatter(Locale.ENGLISH);
        DateTimeFormatter formater = DateTimeFormatter.ofPattern("dd-MM-yyyy");
        dateS = addFullYear(dateS);
        LocalDate locDate = LocalDate.parse(dateS, parser);
        return locDate.format(formater);
    }

    private static String addFullYear(String dateS) {
        String[] dateArr = dateS.split("-");
        dateArr[2] = "20" + dateArr[2];
        return dateArr[0] + "-" + dateArr[1] + "-" + dateArr[2];
    }
}
