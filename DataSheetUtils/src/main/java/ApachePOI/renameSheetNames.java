package ApachePOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileSystemView;
import java.util.*;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


public class renameSheetNames {

    public static int intFileCount = 0;
    public static ArrayList<String> objArrayList = new ArrayList<String>();
    public static LinkedHashMap<String, LinkedHashMap<String, String>> testCaseVarMap = new LinkedHashMap<String, LinkedHashMap<String,String>>();
    public static String strMasterSheetName = "";

    public static void main(String[] args) throws IOException {
        long startTime = System.currentTimeMillis();
        renameSheetNames objRenameSheetNames = new renameSheetNames();
        //String strFolderPath = "C:\\quicktestpro\\DataFiles_Git\\ticketing\\Datafiles\\Ticketing\\GDS Ticketing";
       // String strFolderPath = "C:\\quicktestpro\\DataFiles_Git\\ticketing\\Datafiles\\Ticketing\\AREX";
       String strFolderPath = "C:\\GitCat\\testStandard\\ACS";
        //String strFolderPath = "C:\\quicktestpro\\DataFiles_Git\\ticketing\\Datafiles\\Ticketing\\LATAM";
       //String strFolderPath = objRenameSheetNames.chooseFolder();
        //String strFolderPath = "C:\\Users\\sg0219029\\Desktop\\ZTPF_Function_Run";



        if (strFolderPath.contains("No File/Folder Selected")) {
            System.out.println("No File/Folder Selected");
        }
        else
        {
            objRenameSheetNames.listFiles(strFolderPath);
           System.out.println("No. Of XLS Files in given folder: "+strFolderPath +" are " +intFileCount);

           //check if the .xls file count > 0
           if (intFileCount > 0)
           {
             /*  objRenameSheetNames.readDataFromExcel();
               System.out.println("Reading from Excel files done!!!");
               File objFolder = new File(strFolderPath);
               strMasterSheetName = objFolder.getName();
               objRenameSheetNames.writeDataToExcel(); */
               objRenameSheetNames.updateSheetName();
               System.out.println("Update Sheet Name Done");


           }
        }
        long endTime = System.currentTimeMillis();
        System.out.print("Time taken for this transaction: "+(endTime - startTime) / 1000d + " seconds");
    }



    /*method to return the selected folder path */
    public String chooseFolder()
    {
        JFileChooser objJFileChooser = new JFileChooser();
        objJFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        //objJFileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        objJFileChooser.setCurrentDirectory(FileSystemView.getFileSystemView().getHomeDirectory());

        int returnValue = objJFileChooser.showOpenDialog(null);

        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File objSelectedFile = objJFileChooser.getSelectedFile();
            return(objSelectedFile.getAbsolutePath());

        }
        else {
            return "No File/Folder Selected";
        }
    }


    /* To get all xls files recursively */
    public void listFiles(String strPath)
    {
        File objFolder = new File(strPath);
        File[] objFiles = objFolder.listFiles();

        String strFileExtension = "";
        for (File objFile : objFiles)
        {
            if (objFile.isFile()) {
                String strFileName = objFile.getName();
                String strFilePath = objFile.getAbsolutePath();
                if ( !(strFilePath.toLowerCase().contains("obsolete") || strFilePath.toLowerCase().contains("reusables"))) {

                        if (strFileName.contains(".") && strFileName.lastIndexOf(".") != 0) {
                            strFileExtension = strFileName.substring(strFileName.lastIndexOf(".") + 1);
                            if (strFileExtension.equalsIgnoreCase("xls")) {
                                //objDictionary.put(strFileName, strFilePath);
                                objArrayList.add(strFilePath);
                                intFileCount++;
                            }
                        }
                }
            }
            else if (objFile.isDirectory())
            {
                listFiles(objFile.getAbsolutePath());
            }
        }
    }

/* To read data from each excel file*/
    public  void readDataFromExcel()
    {
        for (String strXLSpath : objArrayList) {
            //System.out.println(strXLSpath);
            try {
                File objFile = new File(strXLSpath);
                String strTestName = objFile.getName();
                FileInputStream objFileInputStream = new FileInputStream(new File(strXLSpath));
                //Get the workbook instance for XLS file
                HSSFWorkbook objWorkbook = new HSSFWorkbook(objFileInputStream);

                //check if no.of sheets in xls file is > 1
                if (objWorkbook.getNumberOfSheets() > 1) {

                    //Get 2nd sheet from the workbook
                    HSSFSheet objSheet = objWorkbook.getSheetAt(1);
                    int intRowStart = objSheet.getFirstRowNum();
                    int intRowEnd = objSheet.getLastRowNum();
                    for(int intRowCounter = intRowStart+1; intRowCounter<=intRowEnd;intRowCounter++)
                    {
                        LinkedHashMap<String, String> variableMap = new LinkedHashMap<String, String>();
                        Row objFirstRow = objSheet.getRow(0);
                        Row objRow = objSheet.getRow(intRowCounter);
                        if(objRow == null || objFirstRow == null)
                        {
                            break;
                        }
                        int intCellStart = objFirstRow.getFirstCellNum();
                        int intCellEnd = objFirstRow.getLastCellNum();
                        boolean blnAddToMap = true;
                        for(int intCellCounter = intCellStart;intCellCounter < intCellEnd;intCellCounter++)
                        {
                            //NUMERIC = 997
                            //STRING = Y
                            //STRING   =  '5.00
                            //FORMULA   CONCATENATE("12",RIGHT(YEAR(TODAY()),2)+1)
                            //BLANK =
                            Cell objFirstRowCell = objFirstRow.getCell(intCellCounter);
                            Cell objCell = objRow.getCell(intCellCounter);
                            if((objCell == null || objCell.toString() == "") && intCellCounter == intCellStart)
                            {
                                blnAddToMap = false;
                                break;
                            }
                            else if(objFirstRowCell == null)
                            {
                                if(intCellCounter == intCellStart) {
                                    blnAddToMap = false;
                                }

                                break;
                            }
                            Enum cellType = CellType.BLANK;
                            String strVariableValue = "";
                            if(objCell!=null) {
                                cellType = objCell.getCellTypeEnum();
                            }

                            if (!(objFirstRowCell.getCellTypeEnum() == CellType.BLANK))
                            {
                                String strVariable = objFirstRowCell.toString();

                                if(!(cellType == CellType.BLANK))
                                {
                                    DataFormatter dataFormatter = new DataFormatter();
                                    strVariableValue = dataFormatter.formatCellValue(objCell);
                                    //strVariableValue = objCell.toString();
                                }

                               // String strVariableValue = objCell.toString();
                                if(cellType == CellType.FORMULA) {
                                    strVariableValue = "=" + strVariableValue;
                                }
                                variableMap.put(strVariable, strVariableValue);
                            }
                            else {
                                if(intCellCounter == intCellStart) {
                                    blnAddToMap = false;
                                }
                                break;
                            }
                        } //loop to read from cell in a row ends here

                        if(!blnAddToMap) {
                            break;
                        }
                        String strTestNameInSheet = strTestName;
                        if (intRowCounter > 1) {
                            strTestNameInSheet = strTestName + "_DataSet_#"+intRowCounter;
                        }
                        testCaseVarMap.put(strTestNameInSheet,variableMap);
                    }//loop to read from rows in a sheet ends here
                }
            } //try ends here
            catch(FileNotFoundException e){
                e.printStackTrace();
            }
            catch(IOException e){
                e.printStackTrace();
            }
            catch(NullPointerException e){
                e.printStackTrace();
                System.out.println("issue at: "+strXLSpath);
            }
        } //for loop which iterate through each excel file ends here
    }



    public void writeDataToExcel()
    {
        /* creating a new master excel workbook*/
        try{
            String strOutputFilePath = FileSystemView.getFileSystemView().getHomeDirectory().toString();
            strOutputFilePath = strOutputFilePath + "\\"+ strMasterSheetName +".xlsx" ;
            System.out.println("Writing to Path: " +strOutputFilePath);
            FileOutputStream objFileOut = new FileOutputStream(strOutputFilePath);
           XSSFWorkbook objNewWorkBook = new XSSFWorkbook();
            //HSSFWorkbook objNewWorkBook = new HSSFWorkbook();
            XSSFSheet objNewSheet = objNewWorkBook.createSheet(strMasterSheetName);
            //HSSFSheet objNewSheet = objNewWorkBook.createSheet(strMasterSheetName);
            CellStyle objStyle = objNewWorkBook.createCellStyle(); //Create new style
            objStyle.setWrapText(true); //Set wordwrap
            Row objFirstRow = objNewSheet.createRow(0);
            Cell objFirstRowCell = objFirstRow.createCell(0);
            objFirstRowCell.setCellValue("TestCase Name");
            int intVarCount = 0;
            int intRowCount = 0;
            LinkedHashMap<String, Integer> masterVariableList = new LinkedHashMap<String, Integer>();

        // loop over the set using an entry set
            for( Map.Entry<String,LinkedHashMap<String, String>> testCaseVarMapEntry : testCaseVarMap.entrySet())
            {
                String strTestName = testCaseVarMapEntry.getKey();
                Pattern objPattern = Pattern.compile("_DataSet_#\\d{1,}");
                Matcher objMatch = objPattern.matcher(strTestName);
                strTestName =  objMatch.replaceAll("");
               /* if(strTestName.matches("_DataSet_#\\d{1,}"))
                {
                    strTestName = strTestName.replaceAll("_DataSet_#\\d{1,}", "");
                    System.out.println("WithDataSet: "+strTestName);
                }*/
                intRowCount++;
                Row objRow = objNewSheet.createRow(intRowCount);
                objRow.createCell(0).setCellValue(strTestName);
                LinkedHashMap<String, String> variableMap = testCaseVarMapEntry.getValue();
               // List<String>value = testCaseVarMapkey.getValue();
                //Set variableMapEntrySet = variableMap.entrySet();

                for( Map.Entry<String,String> variableMapEntrySet : variableMap.entrySet())
                {
                    String strVariableName = variableMapEntrySet.getKey();
                    String strVariableValue = variableMapEntrySet.getValue();

                    if (!masterVariableList.containsKey(strVariableName))
                    {
                        intVarCount++;
                        masterVariableList.put(strVariableName, intVarCount);
                        objFirstRow.createCell(intVarCount).setCellValue(strVariableName);

                    }
                    int intCellPos = (int) masterVariableList.get(strVariableName);
                    Cell objCell = objRow.createCell(intCellPos);

                   // objCell.setCellStyle(objStyle); //Apply style to cell
                    if(strVariableValue.startsWith("="))
                    {
                        strVariableValue = strVariableValue.replace("=","");
                        objCell.setCellFormula(strVariableValue);
                    }
                   /* else if(StringUtils.isNumeric(strVariableValue))
                    {
                        objCell.setCellValue(strVariableValue);
                        objCell.setCellType(CellType.NUMERIC);
                    }*/
                    else
                    {
                        objCell.setCellValue(strVariableValue);
                    }
                    /*if(StringUtils.isNumeric(strVariableValue))
                    {
                        objRow.createCell(intCellPos).setCellValue(strVariableValue);
                        objRow.getCell(intCellPos).setCellType(Cell.CELL_TYPE_NUMERIC);
                        objRow.getCell(intCellPos)
                    }*/
                    //objRow.createCell(intCellPos).setCellValue(strVariableValue);


                }
            }

            objNewWorkBook.write(objFileOut);
            objFileOut.close();
        } catch (IOException e) {
            e.printStackTrace(); }

        System.out.println("Write to excel Completed!!!");
    }



    public  void updateSheetName()
    {
        for (String strXLSpath : objArrayList) {
            //System.out.println(strXLSpath);
            try {
                File objFile = new File(strXLSpath);
                String strTestName = objFile.getName();
                FileInputStream objFileInputStream = new FileInputStream(new File(strXLSpath));
                //Get the workbook instance for XLS file
                HSSFWorkbook objWorkbook = new HSSFWorkbook(objFileInputStream);

                //check if no.of sheets in xls file is > 1
                if (objWorkbook.getNumberOfSheets() > 1) {
                    FileOutputStream objFileOut = new FileOutputStream(strXLSpath);
                    objWorkbook.setSheetName(0, "Flow");
                    objWorkbook.setSheetName(1, "Data");
                    objWorkbook.write(objFileOut);
                    objFileOut.close();
                                //break;
                            }

            } //try ends here
            catch(FileNotFoundException e){
                e.printStackTrace();
            }
            catch(IOException e){
                e.printStackTrace();
            }
            catch(NullPointerException e){
                e.printStackTrace();
                System.out.println("issue at: "+strXLSpath);
            }
        } //for loop which iterate through each excel file ends here
    }


}


