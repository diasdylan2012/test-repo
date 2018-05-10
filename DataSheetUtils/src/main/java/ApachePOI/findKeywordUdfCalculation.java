package ApachePOI;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import java.io.*;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class findKeywordUdfCalculation {

    public void lookThroughDir(File directory) throws Exception {
        File[] folder = directory.listFiles();
        for(File file : folder) {
            if(file.isFile())    {
                if (file.getName().contains(".xls"))
                    findKeywordInFile(file);
            }
            else    {
                if(!(file.getPath().contains("Reusables")))
                lookThroughDir(file);
            }
        }
    }

    private void findKeywordInFile(File file) throws IOException {
        InputStream ExcelFileToRead = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileToRead);
        HSSFSheet flowSheet = workbook.getSheetAt(0);
        HSSFRow row,row1=null,row2=null,row3=null,row4=null,row5=null;
        HSSFCell cell,inputCell,expectedCell = null;
        Pattern p = null;
        boolean testcaseStartFound=false;
        int keywordCol=0,keywordRow=0,count=0;

        Iterator rows = flowSheet.rowIterator();

        //testcase start
        while (rows.hasNext())  {
            row=(HSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()) {
                cell=(HSSFCell) cells.next();
                if(cell.toString().contains("Keywords"))    {
                    keywordRow=cell.getRowIndex();
                    keywordCol = cell.getColumnIndex();
                    testcaseStartFound=true;
                    break;
                }
            }

        }

        if (testcaseStartFound) {
            p = Pattern.compile("(?i)UdfAirthmaticCalculation$");
            Matcher m;

            for (int i=0;i<flowSheet.getLastRowNum();i++)   {
                row=flowSheet.getRow(i);
                if (row.getCell(keywordCol) !=null && row.getCell(keywordCol).toString().trim().equalsIgnoreCase("Testcase End"))
                    break;
                cell=row.getCell(keywordCol);
                inputCell=row.getCell(keywordCol+1);
                expectedCell=row.getCell(keywordCol+2);
                if (cell==null||cell.getCellTypeEnum()== CellType.BLANK)
                    continue;
                m = p.matcher(cell.toString());
                while(m.find()) {
                    if(inputCell != null || inputCell.getCellTypeEnum() != CellType.BLANK)   {
                        System.out.println(file.getAbsolutePath()+"%" +cell.toString() +"%"+inputCell.toString()+"%"+expectedCell.toString());
                    }
                }
            }

        }


    }

    public static void main(String[] args) throws Exception {
        String path = "C:\\GitCat\\Product_Areas\\Ticketing\\Datafiles\\Ticketing\\Airline Ticketing";
        File directory = new File(path);
        findKeywordUdfCalculation obj = new findKeywordUdfCalculation();
        obj.lookThroughDir(directory);
    }

}
