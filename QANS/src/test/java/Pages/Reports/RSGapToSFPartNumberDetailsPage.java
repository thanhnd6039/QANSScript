package Pages.Reports;

import Helpers.Manager.FileReaderManager;
import org.junit.Assert;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import java.util.ArrayList;
import java.util.List;

public class RSGapToSFPartNumberDetailsPage extends RSCommonPage {
    private WebDriver driver;
    @FindBy(xpath = "//*/div[contains(text(),'GAP to SF - Part Number Detail')]")
    private WebElement txtTitle;
    public RSGapToSFPartNumberDetailsPage(WebDriver driver){
        super(driver);
        PageFactory.initElements(driver, this);
        this.driver = driver;
    }
    public void shouldSeeTitle(String title){
        Assert.assertTrue(waitForElementVisibility(txtTitle));
        String actualTitle = getTextFromElement(txtTitle);
        Assert.assertEquals(title, actualTitle);
    }
    public void getOEMGroupByMainSaleRep(){
        String vtOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\VTOEMGroup.xlsx";
        String allOEMWithOEMGroupFilePath = "C:\\CucumberFramework\\Downloads\\AllOEMwithOEMGroup.xlsx";
        List<Object[]> dataOfVTOEMGroup = new ArrayList<>();
        List<Object[]> dataOfAllOEMWithOEMGroup = new ArrayList<>();
        List<Object[]> dataOfOEMGroupByMainSaleRep = new ArrayList<>();
        Object[] colsArrOfOEMGroupByMainSaleRep = new Object[2];
        dataOfVTOEMGroup = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(vtOEMGroupFilePath, 0, 1, 0);
        dataOfAllOEMWithOEMGroup = FileReaderManager.getInstance().getExcelReader().readDataFromExcel(allOEMWithOEMGroupFilePath, 0, 1, 0);
        for (int dataVTOEMGroupIndex = 0; dataVTOEMGroupIndex < dataOfVTOEMGroup.size(); dataVTOEMGroupIndex++){
//          The OEMGroup from VTOEMGroup file
            String oemGroup = dataOfVTOEMGroup.get(dataVTOEMGroupIndex)[1].toString().trim();
//          The SaleRep from VTOEMGroup file
            String saleRepOfOEMGroup = dataOfVTOEMGroup.get(dataVTOEMGroupIndex)[4].toString().trim();
            if (saleRepOfOEMGroup.isEmpty()){
                for (int dataOfAllOEMWithOEMGroupIndex = 0; dataOfAllOEMWithOEMGroupIndex < dataOfAllOEMWithOEMGroup.size(); dataOfAllOEMWithOEMGroupIndex++){
                    String cusName =  dataOfAllOEMWithOEMGroup.get(dataOfAllOEMWithOEMGroupIndex)[2].toString().trim();
                    String saleRepOfCus = dataOfAllOEMWithOEMGroup.get(dataOfAllOEMWithOEMGroupIndex)[2].toString().trim();
                    if (oemGroup.equalsIgnoreCase(cusName)){
                        colsArrOfOEMGroupByMainSaleRep[0] = oemGroup;
                        colsArrOfOEMGroupByMainSaleRep[1] = saleRepOfCus;
                        break;
                    }
                }
            }
            else {
                colsArrOfOEMGroupByMainSaleRep[0] = oemGroup;
                colsArrOfOEMGroupByMainSaleRep[1] = saleRepOfOEMGroup;
            }
            dataOfOEMGroupByMainSaleRep.add(colsArrOfOEMGroupByMainSaleRep);
        }
//        for (int index = 0; index < dataOEMGroupByMainSaleRep.size(); index++){
//            String oemGroupByMainSaleRep = dataOEMGroupByMainSaleRep.get(index)[0].toString().trim();
//            String saleRep = dataOEMGroupByMainSaleRep.get(index)[1].toString().trim();
//            System.out.println(String.format("OEM: %s, Sale: %s", oemGroupByMainSaleRep, saleRep));
//        }
    }
    public void getSourceDataForGapToSFPNDetailReport(){

    }




}
