/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package parcosimportsys;

import com.opencsv.CSVReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author narayan
 */
public class CatlogMatcherNew {
    public static void main(String args[]){
        FileInputStream fis = null;
            int globalcount=1;
            HashSet my_hashset = new HashSet();
            try {
                XSSFWorkbook writerworkbook = new XSSFWorkbook();
                
                XSSFSheet writersheet = writerworkbook.createSheet("parcosout");

                
                  File myFile = new File("/home/narayan/Downloads/parcosam.xlsx");
            fis = new FileInputStream(myFile);
            // Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
            // Return first sheet from the XLSX workbook
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            
                // Creating an empty TreeMap of string and Object][]
                // type
                Object[][] datawriter = {{ "sku"}};
                int rowNum = 0;
                for (Object[] datatype : datawriter) {
                    Row rowan = writersheet.createRow(rowNum++);
                    int colNum = 0;
                    for (Object field : datatype) {
                        Cell cellan = rowan.createCell(colNum++);
                        if (field instanceof String) {
                            cellan.setCellValue((String) field);
                        } else if (field instanceof Integer) {
                            cellan.setCellValue((Integer) field);
                        }
                    }
                }

                

                
                // Create an object of filereader
                // class with CSV file as a parameter.
                String file="/home/narayan/Downloads/catlog.csv";
                FileReader filereader = new FileReader(file);

                // create csvReader object passing
                // file reader as a parameter
                CSVReader csvReader = new CSVReader(filereader);
                String[] nextRecord;

                ArrayList<String> armos=new ArrayList<String>();
                // we are going to read data line by line
                while ((nextRecord = csvReader.readNext()) != null) {
                      armos.add(nextRecord[0]);
//                    for (String cell : nextRecord) {
//                        //System.out.print(cell + "\t");
//                    }
                    
                   // System.out.println(nextRecord[0]);
              

                    
                }
                
                
                
                int rowz=mySheet.getPhysicalNumberOfRows();
                
                System.out.println("One to One Matches:");
                for(int iz=0;iz<rowz;iz++) {
                    Row rowerz=mySheet.getRow(iz);
                    Cell cz=rowerz.getCell(0);
                    for(int ix=0;ix<armos.size();ix++) {
                        if(cz.toString().equals(armos.get(ix))) {
                            System.out.println(cz.toString());
                        }
                        
                    }
                    
                }
                
                
                
                System.out.println("Matches with trailing zeros:");
                for(int iz=1;iz<rowz;iz++) {
                    Row rowerz=mySheet.getRow(iz);
                    Cell cz=rowerz.getCell(0);
                    for(int ix=0;ix<armos.size();ix++) {
                        if(cz.toString().length()>1){
                            if(cz.toString().substring(1).equals(armos.get(ix))) {
                                System.out.println(cz.toString());
                            }
                        }
                        
                    }
                    
                }
                
                
                
                System.out.println("UnMatches:");
                for(int iz=1;iz<rowz;iz++) {
                    Row rowerz=mySheet.getRow(iz);
                    Cell cz=rowerz.getCell(0);
                    for(int ix=0;ix<armos.size();ix++) {
                        if(cz.toString().length()>1){
                            if(!cz.toString().equals(armos.get(ix))) {
                                System.out.println(armos.get(ix));
                            }
                        }
                        
                    }
                    
                }
                
                
                
                
                
        try (FileOutputStream outputStream = new FileOutputStream("/home/narayan/ParcosMissingSku.xlsx")) {
            writerworkbook.write(outputStream);
        }
        
        System.out.println("\n\n\n\nValues are as follows:");
        Iterator itm=my_hashset.iterator();
        
        Object[] arrox=my_hashset.toArray();
        
        for(int ix=0;ix<arrox.length;ix++){
            System.out.println(arrox[ix].toString());
            Row rowexl = writersheet.createRow(globalcount);
            Cell cellsku=rowexl.createCell(0);
            cellsku.setCellValue(arrox[ix].toString());
            globalcount=globalcount+1;
        }
                
                
            }
            catch (Exception e) {
                e.printStackTrace();
            }
        
        
    }
}
