package ApachePOI;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/*
 *Created by SG0227825 on 23/03/2018
 */

public class findUnusedVariables {

    float totalcount;
    public void lookThroughDir(File directory) throws Exception {
        File[] folder = directory.listFiles();
        for(File file : folder) {
            if(file.isFile())    {
                if (!isIgnoredFile(file) && file.getName().contains(".xls"))
                findUnusedVars(file);
            }
            else    {
                if(!(file.getPath().contains("Reusables")))
                    lookThroughDir(file);
            }
        }
    }

    private boolean isIgnoredFile(File file) {
        if (       file.getName().contains("Airports_Master.xls")  //master sheet
                || file.getName().contains("Dsh_Sanity_CheckEnvironment.xls")  //blank data tab
                || file.getName().contains("Data.xls")  //like master sheet
                || file.getName().contains("Data_DMS.xls")  //like master sheet
                || file.getName().contains("Data_EMD_A.xls")  //like master sheet
                || file.getName().contains("Date_CityPairs.xls")  //city pair sheet
                || file.getName().contains("ChkAvlTik.xls")  //only data tab
                || file.getName().contains("IssueTicket.xls")  //only data tab
                || file.getName().contains("ASR_CDR_Master.xls")  //like master sheet
                || file.getName().contains("Dsh_ASR_price.xls")  //blank data tab
                || file.getName().contains("Dsh_DMS_presteps_for_CDR.xls")  //blank data tab
                || file.getName().contains("ASR_CR_Master.xls")  //like master sheet
                || file.getName().contains("ASR_LAN255_Master.xls")  //like master sheet
                || file.getName().contains("Dsh_LANFF06_PreSteps_Data.xls")  //blank data tab
                || file.getName().contains("ASR_OFFLINE_SUITE_Master.xls")  //like master sheet
                || file.getName().contains("Dsh_ASR_presteps_")  //blank data tab
                || file.getName().contains("Dsh_ASR_verify.xls")  //blank data tab
                || file.getName().contains("ASR_T2_EMD_Master.xls")  //like master sheet
                ) {
            return true;
        }
        else
            return false;
    }

    private void findUnusedVars(File file) throws IOException {

        String name = file.getName();
        String path = file.getAbsolutePath().replace("\\"+name,"");
//        System.out.println(path+"\\"+name);

        InputStream ExcelFileToRead = new FileInputStream(file);

        HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileToRead);
        List<String> dataList = new ArrayList<String>(), flowList = new ArrayList<String>();
        HSSFSheet flowSheet = workbook.getSheetAt(0);   //Flow Tab
        HSSFSheet dataSheet = workbook.getSheetAt(1);   //Data Tab
        HSSFRow row;
        HSSFCell cell,keywordCell;

        int keywordCol=0,keywordRow=0;

        //find rowIndex & colIndex for Testcase start
        Iterator rows = flowSheet.rowIterator();
        while (rows.hasNext())  {
            boolean testcaseStartFound=false;
            row=(HSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()) {
                cell=(HSSFCell) cells.next();
                if(cell.toString().contains("Keywords"))    {
                    keywordCol = cell.getColumnIndex();
                    keywordRow = cell.getRowIndex();
                    testcaseStartFound=true;
                    break;
                }
            }
            if (testcaseStartFound)
                break;
        }

        int inputCol=keywordCol+1,expectedCol=keywordCol+2,firstRow=keywordRow+1;
        boolean callReusablesFound = false;
        String unusedVars="";

        //create list for data tab
        row = dataSheet.getRow(0);
        for(int i=0;i<row.getLastCellNum();i++) {
            cell = row.getCell(i);
            if (cell==null || cell.getCellTypeEnum() == CellType.BLANK)
                continue;
            dataList.add(cell.toString().trim());
        }

        //reading nested variables from data tab
        row = dataSheet.getRow(1);
        for (int i=0;i<row.getLastCellNum();i++)    {
            cell = row.getCell(i);
            if (cell==null || cell.getCellTypeEnum() == CellType.BLANK )
                continue;
            Pattern p = Pattern.compile("\\[.*?\\]");
            Matcher m;
            m = p.matcher(cell.toString());
            while (m.find())    {
                flowList.add(m.group(0).toString().toUpperCase().replaceAll("\\[","").replaceAll("\\]",""));
            }
        }

        //reading through Flow tab.
        Pattern p = Pattern.compile("\\[.*?\\]");
        Matcher m;
        for (int i=firstRow;i<flowSheet.getLastRowNum();i++)    {
            row = flowSheet.getRow(i);
            keywordCell = row.getCell(keywordCol);

            if (keywordCell!=null && keywordCell.toString().trim().equalsIgnoreCase("Testcase End"))
                break;
            if (keywordCell!=null && keywordCell.toString().trim().equalsIgnoreCase("UDFCALLREUSABLES"))    {   //UDFCALLREUSABLES keyword - if true - make no changes to script
                callReusablesFound = true;
                break;
            }
            //add to list for Flow tab
            for (int j=inputCol;j<=expectedCol;j++)  {
                cell = row.getCell(j);
                if (cell==null || cell.getCellTypeEnum() == CellType.BLANK)
                    continue;
                m = p.matcher(cell.toString());
                while (m.find())    {
                    flowList.add(m.group(0).toString().toUpperCase().replaceAll("\\[","").replaceAll("\\]",""));
                }
            }
        }

        if(!callReusablesFound)  {  //printing required values
            for(int i=0;i<dataList.size();i++)  {
                if(flowList.contains(dataList.get(i).toUpperCase()))
                    continue;
                else    {
                    unusedVars = unusedVars+"|"+dataList.get(i).toString();
                    totalcount++;
                }
            }
        }

        if(unusedVars!="") {
            unusedVars=unusedVars.substring(1); //remove first | of unusedVars
            System.out.println(name+"%"+path+"%"+unusedVars);
//            System.out.println(totalcount);
        }
    }


    public static void main(String[] args) throws Exception {
        String path = "C:\\GitCat\\testStandard\\PRS";
        File directory = new File(path);
        findUnusedVariables obj = new findUnusedVariables();
        obj.lookThroughDir(directory);

    }
}
