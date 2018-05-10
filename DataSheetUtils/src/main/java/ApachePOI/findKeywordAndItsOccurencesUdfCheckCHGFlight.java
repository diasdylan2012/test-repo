package ApachePOI;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.io.*;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class findKeywordAndItsOccurencesUdfCheckCHGFlight {

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

    private boolean udfCallReusablesFound(HSSFSheet flowSheet, int keywordCol) {
        boolean found= false;
        Iterator rows = flowSheet.rowIterator();
        while (rows.hasNext())  {
            HSSFRow row = (HSSFRow) rows.next();
            HSSFCell cell= row.getCell(keywordCol);
            if (cell==null)
                continue;
            if (cell.toString().trim().equalsIgnoreCase("Testcase End"))
                break;
            if (cell.toString().trim().equalsIgnoreCase("UDFCALLREUSABLES")) {
                found=true;
                break;
            }
        }
        return found;
    }

    private void findKeywordInFile(File file) throws IOException {
        InputStream ExcelFileToRead = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileToRead);
        HSSFSheet flowSheet = workbook.getSheetAt(0);
        HSSFRow row,row0=null,row1=null,row2=null,row3=null,row4=null,row5=null;
        HSSFCell cell;
        Pattern p = null;
        boolean testcaseStartFound=false;
        int keywordCol=0,keywordRow=0,count=0;

        Iterator rows = flowSheet.rowIterator();
        //find testcasestart
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

        //look for keyword
        if (testcaseStartFound && !udfCallReusablesFound(flowSheet,keywordCol)) {
            p = Pattern.compile("(?i)UdfCheckCHGFlight$");//UdfCheckCHGFlight
            Matcher m;
            for (int i=0;i<flowSheet.getLastRowNum();i++)   {
                row=flowSheet.getRow(i);
                if (row.getCell(keywordCol) !=null && row.getCell(keywordCol).toString().trim().equalsIgnoreCase("Testcase End"))
                    break;
                cell=row.getCell(keywordCol);
                if (cell==null||cell.getCellTypeEnum()== CellType.BLANK)
                    continue;
                m = p.matcher(cell.toString());
                while(m.find()) {
                    if (count==0)   {
                        row0=flowSheet.getRow(i-1);
                        row1=flowSheet.getRow(i);
                        row2=flowSheet.getRow(i+1);
                        row3=flowSheet.getRow(i+2);
                        row4=flowSheet.getRow(i+3);
                        row5=flowSheet.getRow(i+4);
                    }
                    count++;
                }
            }
//            System.out.println(file.getAbsolutePath());
            if (count==1)   {
                System.out.print("\n"+file.getAbsolutePath()+"%");

                Iterator<Cell> cellItr = row0.iterator();
                while(cellItr.hasNext()){
                    System.out.print(cellItr.next().toString()+"|");
                }

                System.out.print("%");
                cellItr = row1.iterator();
                while(cellItr.hasNext()){
                    System.out.print(cellItr.next().toString()+"|");
                }
                System.out.print("%");
                cellItr = row2.iterator();
                while(cellItr.hasNext()){
                    System.out.print(cellItr.next().toString()+"|");
                }
                System.out.print("%");
                cellItr = row3.iterator();
                while(cellItr.hasNext()){
                    System.out.print(cellItr.next().toString()+"|");
                }
                System.out.print("%");
                cellItr = row4.iterator();
                while(cellItr.hasNext()){
                    System.out.print(cellItr.next().toString()+"|");
                }
                if (row5!=null) {   //sometimes no 5th row
                    System.out.print("%");
                    cellItr = row5.iterator();
                    while(cellItr.hasNext()){
                        System.out.print(cellItr.next().toString()+"|");
                    }
                }

            }
        }

    }


    public static void main(String[] args) throws Exception {
        String path = "C:\\GitCat\\testStandard\\ACS";
        File directory = new File(path);
        findKeywordAndItsOccurencesUdfCheckCHGFlight obj = new findKeywordAndItsOccurencesUdfCheckCHGFlight();
        obj.lookThroughDir(directory);
    }
}
