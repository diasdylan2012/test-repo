package ApachePOI;

/*
 *Created by SG0227825 on 29/03/2018
 */

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class variableStandardization {

    public void lookThroughDir(File directory) throws Exception {
        File[] folder = directory.listFiles();
        for(File file : folder) {
            if(file.isFile())    {
                if (!isIgnoredFile(file) && file.getName().contains(".xls"))
                    searchVariableSynonym(file);
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

    private boolean isIgnoredFile(File file) {
        if (       file.getName().equals("Airports_Master.xls")  //master sheet
                || file.getName().equals("Dsh_Sanity_CheckEnvironment.xls")  //blank data tab
                || file.getName().equals("Data.xls")  //like master sheet
                || file.getName().equals("Data_DMS.xls")  //like master sheet
                || file.getName().equals("Data_EMD_A.xls")  //like master sheet
                || file.getName().equals("Date_CityPairs.xls")  //city pair sheet
                || file.getName().equals("ChkAvlTik.xls")  //only data tab
                || file.getName().equals("IssueTicket.xls")  //only data tab
                || file.getName().equals("ASR_CDR_Master.xls")  //like master sheet
                || file.getName().equals("Dsh_ASR_price.xls")  //blank data tab
                || file.getName().equals("Dsh_DMS_presteps_for_CDR.xls")  //blank data tab
                || file.getName().equals("ASR_CR_Master.xls")  //like master sheet
                || file.getName().equals("ASR_LAN255_Master.xls")  //like master sheet
                || file.getName().equals("Dsh_LANFF06_PreSteps_Data.xls")  //blank data tab
                || file.getName().equals("ASR_OFFLINE_SUITE_Master.xls")  //like master sheet
                || file.getName().equals("Dsh_ASR_presteps_")  //blank data tab
                || file.getName().equals("Dsh_ASR_verify.xls")  //blank data tab
                || file.getName().equals("ASR_T2_EMD_Master.xls")  //like master sheet
                ) {
            return true;
        }
        else
            return false;
    }

    private void searchVariableSynonym(File file) throws IOException {

        //for each standardized keyword in input file
        String inputFilePath = "C:\\GitCat\\testStandard\\standardization.txt";
        File standardizationInput = new File(inputFilePath);
        try (BufferedReader br = new BufferedReader(new FileReader(standardizationInput))) {
            String line;
            while ((line = br.readLine()) != null) {    //for each line in standardization input file

                String name = file.getName();
                String path = file.getAbsolutePath().replace("\\"+name,"");
//
                InputStream excelFileToRead = new FileInputStream(file);
                HSSFWorkbook workbook = new HSSFWorkbook(excelFileToRead);
                HSSFSheet flowSheet = workbook.getSheetAt(0);   //Flow Tab
                HSSFSheet dataSheet = workbook.getSheetAt(1);   //Data Tab
                HSSFRow row;
                HSSFCell cell,keywordCell;
                List<Integer> dataList = new ArrayList<>();
                String standardVar = null, nonStandardVar;
                List<String> oldVarNames = new ArrayList<>();

                standardVar=line.split(":")[0];
                nonStandardVar=line.split(":")[1];
                for (String s : nonStandardVar.split(",")) {
                    oldVarNames.add(s);
                }

                //find all unstandardized var names
                row = dataSheet.getRow(0);
                for (int i=0;i<row.getLastCellNum();i++)    {
                    cell = row.getCell(i);
                    if (cell==null)
                        continue;
                    if(oldVarNames.contains(cell.toString()))
                        dataList.add(cell.getColumnIndex());
                }

                //
                int keywordCol=0,keywordRow=0;

                //find rowIndex & colIndex for Testcase start
                Iterator rows = flowSheet.rowIterator();
                if (dataList.size()>0)  {
                    boolean testcaseStartFound=false;
                    while (rows.hasNext())  {
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
                }

                if (!udfCallReusablesFound(flowSheet,keywordCol) && dataList.size()>0)  {
                    int inputCol=keywordCol+1,expectedCol=keywordCol+2,firstRow=keywordRow+1;
                    List<String> varNames = new ArrayList<>(),varValues = new ArrayList<>(),NewvarNames = new ArrayList<>();

                    //list of old var names
                    for (int x=0;x<dataList.size();x++)
                        varNames.add(dataSheet.getRow(0).getCell(dataList.get(x)).toString());
                    
                    //list of standardized var names
                    for (int x=0;x<dataList.size();x++) {
                        String newVarName = standardVar + (x + 1);
                        if (x==0)
                            newVarName=newVarName.substring(0,newVarName.length()-1);
                        NewvarNames.add(newVarName);
                    }

                    //keeping track of which standardized var names have been used
                    int[] unused = new int[varNames.size()];
                    for (int x=0;x<unused.length;x++) {
                        String newVarName= standardVar+(x+1);
                        if (x==0)
                            newVarName=newVarName.substring(0,newVarName.length()-1);
                        if (varNames.contains(newVarName))
                            unused[x]=1;
                    }

                    //for each synonym variable
                    for (int i=0;i<varNames.size();i++) {
                        String newVarName;
                        if (!NewvarNames.contains(varNames.get(i))) {
                            newVarName = getUnusedVarName(standardVar,unused);
                            String pattern1 = "(?i)\\[\\b"+varNames.get(i)+"\\b\\]"; //[Partition_Name]
                            String pattern2 = "(?i)\\{\\b"+varNames.get(i)+"\\b\\}"; //{Partition_Name}
                            Pattern p1 = Pattern.compile(pattern1); //[Partition_Name]
                            Pattern p2 = Pattern.compile(pattern2); //{Partition_Name}

                            //go through flow n standardize
                            for (int j=firstRow;j<flowSheet.getLastRowNum();j++)    {
                                row = flowSheet.getRow(j);
                                keywordCell = row.getCell(keywordCol);
                                if (keywordCell!=null && keywordCell.toString().trim().equalsIgnoreCase("Testcase End"))
                                    break;
                                for (int k=inputCol;k<=expectedCol;k++)  {
                                    cell = row.getCell(k);
                                    if (cell==null )
                                        continue;
                                    Matcher m1 = p1.matcher(cell.toString());
                                    while (m1.find())    {
                                        String newValue = cell.toString().replaceAll(pattern1,"["+newVarName+"]");
                                        cell.setCellValue(newValue);
                                    }
                                    Matcher m2 = p2.matcher(cell.toString());
                                    while (m2.find())    {
                                        String newValue = cell.toString().replaceAll(pattern2,"{"+newVarName+"}");
                                        cell.setCellValue(newValue);
                                    }
                                }
                            }
                            //go through data and standardise
                            dataSheet.getRow(0).getCell(dataList.get(i)).setCellValue(newVarName);

                            //go through data values and standardise
                            row = dataSheet.getRow(1);
                            for (int j=0;j<row.getLastCellNum();j++)    {
                                cell=row.getCell(j);
                                if (cell==null )
                                    continue;
                                Matcher m = p1.matcher(cell.toString());
                                while (m.find())    {
                                    String newValue = cell.toString().replaceAll(pattern1,"["+newVarName+"]");
                                    cell.setCellValue(newValue);
                                }
                            }
                        }
                    }

                    int[] duplicateFound = new int[varNames.size()];
                    Arrays.fill(duplicateFound,0);
                    String unusedVars = "";

                    //list of varValues
                    for (int x=0;x<dataList.size();x++) {
                        String val;
                        HSSFRow tempRow = dataSheet.getRow(1);
                        HSSFCell tempCell = tempRow.getCell(dataList.get(x));
                        if (tempCell==null)
                            val = "";
                        else
                            val = tempCell.toString();

                        varValues.add(val);
                    }

                    //update varNames list
                    varNames.clear();
                    for (int x=0;x<dataList.size();x++)
                        varNames.add(dataSheet.getRow(0).getCell(dataList.get(x)).toString());

                    //delete variables with same values
                    for (int i=0;i<varValues.size()-1;i++) {
                        for (int j=i+1;j<varValues.size();j++) {
                            if (varValues.get(i)==varValues.get(j) && duplicateFound[j]==0) {
                                String newVarName = varNames.get(i);
                                String pattern1 = "(?i)\\[\\b"+varNames.get(j)+"\\b\\]"; //[Partition_Name]
                                String pattern2 = "(?i)\\{\\b"+varNames.get(j)+"\\b\\}"; //{Partition_Name}
                                Pattern p1 = Pattern.compile(pattern1);
                                Pattern p2 = Pattern.compile(pattern2);
                                
                                //delete from flow
                                for (int k=firstRow;k<flowSheet.getLastRowNum();k++)    {
                                    row = flowSheet.getRow(k);
                                    keywordCell = row.getCell(keywordCol);
                                    if (keywordCell!=null && keywordCell.toString().trim().equalsIgnoreCase("Testcase End"))
                                        break;
                                    for (int l=inputCol;l<=expectedCol;l++)  {
                                        cell = row.getCell(l);
                                        if (cell==null )
                                            continue;
                                        Matcher m1 = p1.matcher(cell.toString());
                                        while (m1.find())    {
                                            String newValue = cell.toString().replaceAll(pattern1,"["+newVarName+"]");
                                            cell.setCellValue(newValue);
                                        }
                                        Matcher m2 = p2.matcher(cell.toString());
                                        while (m2.find())    {
                                            String newValue = cell.toString().replaceAll(pattern2,"{"+newVarName+"}");
                                            cell.setCellValue(newValue);
                                        }
                                    }
                                }

                                //replace data values
                                row = dataSheet.getRow(1);
                                for (int k=0;k<row.getLastCellNum();k++)    {
                                    cell=row.getCell(k);
                                    if (cell==null )
                                        continue;
                                    Matcher m = p1.matcher(cell.toString());
                                    while (m.find())    {
                                        String newValue = cell.toString().replaceAll(pattern1,"["+newVarName+"]");
                                        cell.setCellValue(newValue);
                                    }
                                }

                                unusedVars = unusedVars+"|"+varNames.get(j);
                                duplicateFound[i]=1;
                                duplicateFound[j]=1;
                            }

                        }
                    }

                    if(unusedVars!="") {
                        unusedVars=unusedVars.substring(1); //remove first | of unusedVars
                        System.out.println(name+"%"+path+"%"+unusedVars);
                    }

                    excelFileToRead.close();
                    OutputStream excelFileToWrite = new FileOutputStream(file);
                    workbook.write(excelFileToWrite);
                    excelFileToWrite.close();
                }
            }
        }
//        System.out.println(file.getAbsolutePath());
    }

    private String getUnusedVarName(String standardVar, int[] unused) {
        String newVarName = null;
        for (int i=0;i<unused.length;i++)   {
            if (unused[i]==0)   {   //if theres an unused variable name
                newVarName= standardVar+(i+1);
                if (i==0)
                    newVarName=newVarName.substring(0,newVarName.length()-1);
                unused[i]=1;
                break;
            }
        }
        return newVarName;
    }


    public static void main(String[] args) throws Exception {
        String path = "C:\\GitCat\\testStandard\\BMAS"; //C:\GitCat\tpf_regression\Datafiles\Ticketing\Airline Ticketing  //C:\GitCat\testStandard
        File directory = new File(path);
        variableStandardization obj = new variableStandardization();
        obj.lookThroughDir(directory);
    }
}
