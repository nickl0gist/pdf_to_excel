import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

/**
 * Created on 10.03.2021
 * @author Mykola Horkov
 * mykola.horkov@gmail.com
 */
@Setter
@Getter
@ToString
public class Manifest {
    private String dateCollect = "";
    private String dateDeliver = "";
    private String plant = "";
    private int invertedDate = 0;
    private String manifest = "";
    private String supplier = "";
    private int palletQty = 0;
    private double qbm = 0.0;
    private double weight = 0.0;

    public Manifest(String dateCollect, String dateDeliver, String plant, int invertedDate, String manifest,
                    String supplier, int palletQty, double qbm, double weight) {
        this.dateCollect = dateCollect;
        this.dateDeliver = dateDeliver;
        this.plant = plant;
        this.invertedDate = invertedDate;
        this.manifest = manifest;
        this.supplier = supplier;
        this.palletQty = palletQty;
        this.qbm = qbm;
        this.weight = weight;
    }

    public Manifest() {
    }
}
