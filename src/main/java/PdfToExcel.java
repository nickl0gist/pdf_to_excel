import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Created on 10.03.2021
 * @author Mykola Horkov
 * mykola.horkov@gmail.com
 */
public class PdfToExcel {

    private static final String TOTAL_WEIGHT = "TOTAL WEIGHT";
    private static final String COLLECT_FROM = "-------- Collect from: --------";
    private static final String SUPPLIER_CODE = "SUPPLIER CODE";
    private static final String COLLECTION_DATE = "COLLECTION DATE";

    private static ArrayList<Manifest> manifests = new ArrayList<>();

    public static void main(String[] args) throws URISyntaxException {
        String pathToDir = getPath();
        try (Stream<Path> walk = Files.walk(Paths.get(pathToDir))) {
            List<String> result = walk
                    .filter(p -> !Files.isDirectory(p))   // not a directory
                    .map(Path::toString)                  // convert path to string
                    .filter(f -> f.endsWith("pdf"))       // check end with
                    .collect(Collectors.toList());        // collect all matched to a List
            result.forEach(PdfToExcel::readPdf);
        } catch (IOException e) {
            e.printStackTrace();
        }
        ExcelCreator.instanceToExcelFromTemplate(manifests, pathToDir+"\\manifests.xlsx");
    }

    private static String getPath() throws URISyntaxException{
       return new File(PdfToExcel.class.getProtectionDomain().getCodeSource().getLocation()
                .toURI()).getParentFile().getPath();
    }

    public static void readPdf(String pathToPdf) {
        PDFManager pdfManager = new PDFManager();
        pdfManager.setFilePath(pathToPdf);
        try {
            String[] text = pdfManager.toText().split("\\r?\\n");
            manifests.add(readManifest(text));
        } catch (IOException ex) {
            Logger.getLogger(PdfToExcel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private static Manifest readManifest(String[] text) {

        String dateCollect = "";
        String dateDeliver = "";
        String plant = text[1];
        int invertedDate = 0;
        String manifest = "";
        String supplier = "";
        int palletQty = 0;
        double qbm = 0.0;
        double weight = 0.0;

        for (int i = 0; i < text.length; i++) {
            if (text[i].contains(TOTAL_WEIGHT)) {
                manifest = text[i - 1].split(" ")[4];
                String[] cargo = text[i + 1].split(" ");
                weight = Double.parseDouble(cargo[0]);
                qbm = Double.parseDouble(cargo[2]);
                palletQty = Integer.parseInt(cargo[4]);
            } else if (text[i].contains(COLLECT_FROM)) {
                supplier = text[i - 1];
            } else if (text[i].contains(SUPPLIER_CODE)) {
                invertedDate = Integer.parseInt(text[i + 1].split(" ")[1]);
            } else if (text[i].contains(COLLECTION_DATE)) {
                dateCollect = text[i + 1].split(" ")[0];
                dateDeliver = text[i + 1].split(" ")[2];
            }
        }
        return new Manifest(dateCollect, dateDeliver, plant, invertedDate, manifest, supplier, palletQty, qbm, weight);
    }
}
