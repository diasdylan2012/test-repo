package ApachePOI;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class replaceUdfCheckCHGFlight {

    /*
    public void lookThroughDir(File dir) throws IOException {
        File[] folder = dir.listFiles();
        for(File file : folder) {
            if(file.isFile())    {
                if (file.getName().contains(".xls"))
                    replaceKeyword(file);
            }
            else    {
                if(!(file.getPath().contains("Reusables")))
                    lookThroughDir(file);
            }
        }
    } */

    public void readFromListOfFiles (File inputFile) throws IOException {
        try (BufferedReader br = new BufferedReader(new FileReader(inputFile))) {
            String line;
            while ((line = br.readLine()) != null)  {
                File file1= new File(line.trim());
                replaceKeyword(file1);
            }
        }
    }


    public static void removeRow(HSSFSheet sheet, int rowIndex) {
        int lastRowNum=sheet.getLastRowNum();
        if(rowIndex>=0&&rowIndex<lastRowNum){
            sheet.shiftRows(rowIndex+1,lastRowNum, -1);
        }
        if(rowIndex==lastRowNum){
            HSSFRow removingRow=sheet.getRow(rowIndex);
            if(removingRow!=null){
                sheet.removeRow(removingRow);
            }
        }
    }

    private void replaceKeyword(File file) throws IOException {
        InputStream excelRead = new FileInputStream(file);
        HSSFWorkbook wb = new HSSFWorkbook(excelRead);
        HSSFSheet flowSheet = wb.getSheetAt(0);
        HSSFSheet dataSheet = wb.getSheetAt(1);
        List<String> dataList = new ArrayList<String>();
        HSSFRow row = null;
        HSSFCell cell = null;
        int keywordRow=0,keywordCol=0,inputCol=0,expectedCol=0;
        int loc0,loc1,loc2,loc3;
        boolean testcaseStartFound =false,replace =false;

        Iterator rows = flowSheet.rowIterator();
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

        //Data Variable Names
        row = dataSheet.getRow(0);
        for(int i=0;i<row.getLastCellNum();i++) {
            cell = row.getCell(i);
            if (cell==null || cell.getCellTypeEnum() == CellType.BLANK)
                continue;
            dataList.add(cell.toString().trim().toUpperCase());
        }

        if (testcaseStartFound) {
            inputCol=keywordCol+1;
            expectedCol=keywordCol+2;
            Pattern p = Pattern.compile("(?i)UdfCheckCHGFlight$");
            Matcher m;

            for (int i=0;i<flowSheet.getLastRowNum();i++)   {
                loc1=loc2=loc3=0;
                replace=false;
                row=flowSheet.getRow(i);
                if (row.getCell(keywordCol) !=null && row.getCell(keywordCol).toString().trim().equalsIgnoreCase("Testcase End"))
                    break;
                cell=row.getCell(keywordCol);
                if (cell==null||cell.getCellTypeEnum()== CellType.BLANK)
                    continue;
                m = p.matcher(cell.toString());
                while(m.find()) {
                    replace=true;
                    loc1=cell.getRowIndex();

                    if (flowSheet.getRow(loc1-1).getCell(inputCol)!=null && flowSheet.getRow(loc1-1).getCell(inputCol).toString().startsWith("1"))
                        loc0=loc1-1;
                    else if (flowSheet.getRow(loc1-2).getCell(inputCol)!=null && flowSheet.getRow(loc1-2).getCell(inputCol).toString().startsWith("1"))
                        loc0=loc1-2;
                    else
                        loc0=0;

                    if (flowSheet.getRow(loc1+1).getCell(inputCol)!=null && flowSheet.getRow(loc1+1).getCell(inputCol).toString().startsWith("0"))
                        loc2=loc1+1;
                    else if(flowSheet.getRow(loc1+2).getCell(inputCol)!= null && flowSheet.getRow(loc1+2).getCell(inputCol).toString().startsWith("0"))
                        loc2=loc1+2;
                    else    {
                        replace=false;
//                        System.out.println("#");
                        break;
                    }

                    if ((flowSheet.getRow(loc2+1).getCell(keywordCol) != null && flowSheet.getRow(loc2+1).getCell(keywordCol).toString().matches("(?i)UdfRandomName")) || (flowSheet.getRow(loc2+1).getCell(inputCol)!=null && flowSheet.getRow(loc2+1).getCell(inputCol).toString().startsWith("-")))
                        loc3=loc2+1;
                    else if ((flowSheet.getRow(loc2+2).getCell(keywordCol)!=null && flowSheet.getRow(loc2+2).getCell(keywordCol).toString().matches("(?i)UdfRandomName")) || (flowSheet.getRow(loc2+2).getCell(inputCol)!=null && flowSheet.getRow(loc2+2).getCell(inputCol).toString().startsWith("-")))
                        loc3=loc2+2;
                    else    {
                        replace=false;
//                        System.out.println("#");
                        break;
                    }
                    if (replace)    {
                        String newInput="1,";
                        String expectedOutput="";    //,{DepDate}


                        HSSFRow tempRow0 = flowSheet.getRow(loc0);
                        HSSFRow tempRow = flowSheet.getRow(loc1);   //udfcheckchgflight
                        HSSFRow tempRow2 = flowSheet.getRow(loc2);  //0.....sell entry

                        HSSFCell tempCell2 = tempRow2.getCell(inputCol);            //sell entry input
                        HSSFCell tempCell = tempRow.getCell(inputCol);              //keyword input
                        HSSFCell tempCell0 = tempRow0.getCell(inputCol);            //avail entry input
                        HSSFCell tempCell3 = tempRow.getCell(inputCol+1);   //keyword expected

                        Pattern depCity = Pattern.compile("(?i)\\[depcity.*?\\]|\\[Cnt_2_City\\]");
                        Pattern arrCity = Pattern.compile("(?i)\\[arrcity.*?\\]");
                        Pattern depDate = Pattern.compile("(?i)\\[\\w*date\\d*\\]");
                        Pattern COS = Pattern.compile("(?i)\\[COS.*?\\]");
                        Matcher m1;
                        Matcher m2 = depDate.matcher(tempCell0.toString());

                        //DepDate
                        m1 = depDate.matcher(tempCell2.toString());
                        if (m1.find())  {
                            newInput=newInput+m1.group(0).toString()+",";
                            expectedOutput = ",{"+m1.group(0).toString().replaceAll("\\[","").replaceAll("\\]","")+"}";
                        }
                        else if (m2.find() && loc0!=0){
                            newInput=newInput+m2.group(0).toString()+",";
                            expectedOutput = ",{"+m2.group(0).toString().replaceAll("\\[","").replaceAll("\\]","")+"}";
                        }
                        else
                            newInput=newInput+"[noDate],";

                        //DepCity
                        m1 = depCity.matcher(tempCell.toString());
                        if (m1.find())
                            newInput=newInput+m1.group(0).toString();
                        else
                            newInput=newInput+"[noDep],";

                        //ArrCity
                        m1 = arrCity.matcher(tempCell.toString());
                        if (m1.find())
                            newInput=newInput+m1.group(0).toString()+",";
                        else
                            newInput=newInput+"[noArr],";

                        //COS
                        m1 = COS.matcher(tempCell2.toString());
                        if (m1.find())
                            newInput=newInput+m1.group(0).toString().toUpperCase()+",";
                        else if (tempCell2.toString().toUpperCase().contains("01Y1"))
                            newInput=newInput+"Y,";
                        else
                            newInput=newInput+"[noCOS],";

                        //No Of Passengers
                        Pattern NoOfP = Pattern.compile("\\]NN|\\]MM");
                        String temp=tempCell2.toString();
                        m1=NoOfP.matcher(temp);

                        if(m1.find())   {
                            newInput = newInput + temp.split(NoOfP.toString())[1]+",";
                        }   else if (temp.substring(1).startsWith("["))  {
                            String temp2="";
                            temp2 = temp.substring(temp.indexOf("["));
                            temp2 = temp2.substring(0,temp2.indexOf("]")+1);
                            newInput = newInput + temp2+",";
                        }   else    {
                            newInput = newInput + temp.charAt(1)+",";
                        }

                        // Partition
                        newInput = newInput + "[Partition]";

                        //get Flight Number variable name from udfcheckCHGflight
                        temp=tempCell3.toString().split(",")[1].replaceAll("\\{","").replaceAll("\\}","");

                        //check if Flight Number variable name is the same for sell entry
                        if (tempCell2.toString().toUpperCase().contains(temp.toUpperCase()))    {
                            expectedOutput = "{"+temp+"}"+expectedOutput;

//                            System.out.println(file.getName()+"UdfGetAvailabilityAndSellForCheckin \t|\t"+newInput +"\t|\t"+expectedOutput);

                            //code to replace keyword row
                            tempRow.getCell(keywordCol).setCellValue("UdfGetAvailabilityAndSellForCheckin");  //keyword replace
                            tempRow.getCell(inputCol).setCellValue(newInput);   //input_value replace
                            tempRow.getCell(inputCol+1).setCellValue(expectedOutput);

                            //delete avail entry and sell entry
                            removeRow(flowSheet,loc2);
                            removeRow(flowSheet,loc0);


//                            excelRead.close();
//                            OutputStream excelWrite = new FileOutputStream(file);
//                            wb.write(excelWrite);
//                            excelWrite.close();

                            System.out.println(file.getName());
                        } else
//                            System.out.println("Unable to Replaced in: "+file.getName());

                        break;
                    }

                }
            }
        }


    }

    public static void main(String[] args) throws IOException {
        File file = new File("C:\\GitCat\\testStandard\\input1.txt");
        replaceUdfCheckCHGFlight rep1 = new replaceUdfCheckCHGFlight();
        rep1.readFromListOfFiles(file);
    }

}
