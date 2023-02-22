/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package parcosimportsys;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.StringTokenizer;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.StringTokenizer;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author narayan
 */
public class MakeupImportSys2 {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        FileInputStream fis = null;
        ArrayList<String> parentstr=new ArrayList<String>(); 
        
        ArrayList<NodeClassifier> childs=new ArrayList<NodeClassifier>();
        
        ArrayList<Integer> parentrows=new ArrayList<Integer>();
        ArrayList<Integer> childrows=new ArrayList<Integer>();
        ArrayList<Integer> singlerows=new ArrayList<Integer>();
        

        
        ArrayList<NodeClassifier> ncs=new ArrayList<NodeClassifier>();
        
        int row=0;
        int col=0;
        boolean parent=false;
        int parentrownum=0;
        int lastparentid=0;
        int parentcnt=0;
        int globalcount=1;
        boolean children=false;
        boolean singlerow=false;
        XSSFWorkbook writerworkbook = new XSSFWorkbook();
        try {
            // TODO code application logic here
            
        DataFormat fmt = writerworkbook.createDataFormat();
        CellStyle cellStyle = writerworkbook.createCellStyle();            
        cellStyle.setDataFormat(fmt.getFormat("@"));
            // Blank workbook

  
        // Creating a blank Excel sheet
        XSSFSheet writersheet
            = writerworkbook.createSheet("parcossheet");
  
        // Creating an empty TreeMap of string and Object][]
        // type
        Object[][] datawriter = {{ "sku","store_view_code","attribute_set_code","product_type","categories","name","description","meta_title","meta_keyword","meta_description","product_websites","short_description","weight","product_online","tax_class_name","hsn_code","visibility","price","shelf_life","fregrance_color_name","gender","product_contains_liquid","sm2_product_features","dimensions","country","size","unit","ref","brand_name","is_featured","is_bestseller","occasion","scent","manufacturer_detail","packer_detail","manufacturing_name","fragrance_family","fragrence_personality","authenticity","configurable_variations","configurable_variation_labels"}};
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
            
            
            
            File myFile = new File("/home/narayan/Downloads/makupmass.xlsx");
            fis = new FileInputStream(myFile);
            // Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
            // Return first sheet from the XLSX workbook
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);
            
            
            int numrows=mySheet.getPhysicalNumberOfRows();
            
            for(row=0;row<numrows;row++) {
              Row currentRow=mySheet.getRow(row);
                int numcols=currentRow.getPhysicalNumberOfCells();
                //parent=false;
              
                    //for(col=0;col<numcols;col++) {


                        if(currentRow.getCell(1,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().equals("Parent")) {
                            int zenx=row+1;
                            System.out.println("Parent Node:"+zenx);
                            ncs.add(new NodeClassifier(zenx,"Parent"));
                        }
                        else if(currentRow.getCell(1,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().equals("Child")) {
                            ncs.add(new NodeClassifier(row,"Child"));
                        }
                        else if(currentRow.getCell(1,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().equals("Single")) {
                            ncs.add(new NodeClassifier(row,"Single"));
                        }




                    //    System.out.print(currentRow.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)+",");
                   // }
              
             
              
              
              //System.out.println();
            }
            
            
            NodeClassifier temp=null;
            
            for(int i=0;i<ncs.size();i++) {
            Row rowex = writersheet.createRow(globalcount); 
                if(ncs.get(i).getClassifier().equals("Parent")){
                    if(temp!=null) {

                      Row roz=mySheet.getRow(ncs.get(i).getNodenumber());  
                      System.out.println("Node ID:"+temp.getNodenumber()+" Node Classifier:"+temp.getClassifier()+" Value of SKU:"+roz.getCell(0)+"-P"+"  Global Count:"+globalcount);
                      
                      Cell cellsku=rowex.createCell(0);
                      cellsku.setCellValue(roz.getCell(0,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim()+"-P");
                      
                      Cell cellstoreview=rowex.createCell(1);
                      cellstoreview.setCellValue("");
                      
                      Cell cellattribsetcode=rowex.createCell(2);
                      cellattribsetcode.setCellValue(roz.getCell(2,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
                      
                      Cell cellproducttype=rowex.createCell(3);
                      cellproducttype.setCellValue("configurable");
                     
                      Cell cellcategories=rowex.createCell(4);
                      cellcategories.setCellValue("");
                     
                      Cell cellname=rowex.createCell(5);
                      cellname.setCellValue(roz.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
                
                      Cell celldesc=rowex.createCell(6);
                      celldesc.setCellValue(roz.getCell(8,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                      
                      Cell cellmeta=rowex.createCell(7);
                      cellmeta.setCellValue(roz.getCell(15,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
                      
                      Cell cellmetakeywords=rowex.createCell(8);
                      cellmetakeywords.setCellValue(roz.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                      
                      Cell cellmetadesc=rowex.createCell(9);
                      cellmetadesc.setCellValue(roz.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      Cell cellprodweb=rowex.createCell(10);
                      cellprodweb.setCellValue("base");
                      
                      Cell cellshortdesc=rowex.createCell(11);
                      cellshortdesc.setCellValue("");  
                      
                      
                      Cell cellweight=rowex.createCell(12);
                      StringTokenizer stkwt=new StringTokenizer(roz.getCell(10,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString(),".");
                      cellweight.setCellValue(stkwt.nextToken());
                      
                      Cell cellprodonline=rowex.createCell(13);
                      cellprodonline.setCellValue("1");              

                      Cell celltaxclass=rowex.createCell(14);
                      celltaxclass.setCellValue(roz.getCell(12,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellhsn=rowex.createCell(15);
                      cellhsn.setCellStyle(cellStyle);
                      cellhsn.setCellValue(roz.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
                      
                      Cell cellvisibility=rowex.createCell(16);
                      cellvisibility.setCellValue("Catalog, Search");

                      
                      Cell cellprice=rowex.createCell(17);
                      cellprice.setCellValue(roz.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellshelflife=rowex.createCell(18);
                      cellshelflife.setCellValue(roz.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellfragrancecolor=rowex.createCell(19);
                      cellfragrancecolor.setCellValue("Transparent");
                      
                      
                      Cell cellgender=rowex.createCell(20);
                      cellgender.setCellValue(roz.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                      
                      Cell cellcontainsliquid=rowex.createCell(21);
                      cellcontainsliquid.setCellValue(roz.getCell(23,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      
                      Cell sm2prodfeatures=rowex.createCell(22);
                      sm2prodfeatures.setCellValue(roz.getCell(34,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      Cell celldimensions=rowex.createCell(23);
                      celldimensions.setCellValue(roz.getCell(46,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      
                      Cell cellcountry=rowex.createCell(24);
                      cellcountry.setCellValue(roz.getCell(47,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());   
                      
                      
                      Cell cellbotsize=rowex.createCell(25);
                      String poko=roz.getCell(27,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                      StringTokenizer stk=new StringTokenizer(poko,".");
                      cellbotsize.setCellValue(poko);
                      
                      
                      Cell cellunits=rowex.createCell(26);
                      cellunits.setCellValue(roz.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      
                      Cell cellrefs=rowex.createCell(27);
                      StringTokenizer strefs=new StringTokenizer(roz.getCell(53,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString(),".");
                      cellrefs.setCellValue(strefs.nextToken());
                      
                      Cell cellbrands=rowex.createCell(28);
                      cellbrands.setCellValue(roz.getCell(54,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      Cell cellisfeatured=rowex.createCell(29);
                      cellisfeatured.setCellValue((roz.getCell(55,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().equals("Y")?"yes":"no"));

                      
                      Cell cellisbestseller=rowex.createCell(30);
                      cellisbestseller.setCellValue("");
                      
                      Cell celloccasion=rowex.createCell(31);
                      celloccasion.setCellValue(roz.getCell(56,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell cellscent=rowex.createCell(32);
                      cellscent.setCellValue(roz.getCell(57,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      Cell cellmanufacturer=rowex.createCell(33);
                      cellmanufacturer.setCellValue(roz.getCell(58,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell cellpacker=rowex.createCell(34);
                      cellpacker.setCellValue(roz.getCell(59,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell cellmanufname=rowex.createCell(35);
                      cellmanufname.setCellValue(roz.getCell(60,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      Cell cellfragfamily=rowex.createCell(36);
                      cellfragfamily.setCellValue(roz.getCell(20,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      Cell cellfragpersona=rowex.createCell(37);
                      cellfragpersona.setCellValue(roz.getCell(21,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      Cell cellauthenticity=rowex.createCell(38);
                      cellauthenticity.setCellValue(roz.getCell(64,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Row rowold=writersheet.getRow(lastparentid);
                      
                      System.out.println("Old Row number:"+(i-childs.size()));
                      
                      String genrated="";
                      int sizex=(childs.size()-1);
                      for(int chx=0;chx<childs.size();chx++) {
                          Row rowtmp=mySheet.getRow(childs.get(chx).getNodenumber());
                          System.out.println("FOUND THE NOS AS FOLLOWS:"+childs.get(chx).getNodenumber());
                          if(sizex==chx){
                            String pokox=roz.getCell(27,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                            StringTokenizer stkx=new StringTokenizer(pokox,".");
                            //cellbotsize.setCellValue(stkx.nextToken());
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",size="+pokox+"";
                          }
                          else {
                              
                            String pokox=roz.getCell(27,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                            StringTokenizer stkx=new StringTokenizer(pokox,".");
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",size="+pokox+"|";
                          }
                          System.out.println("Generated String as follows.");
                      }
                      
                      Cell configurablevariations=rowold.createCell(39);
                      configurablevariations.setCellValue(genrated);
                      
                      Cell configurablevariationslabel=rowex.createCell(40);
                      configurablevariationslabel.setCellValue("size=Size");
                      
                      
                      childs.clear();
                      
                      lastparentid=globalcount;
//datawriter(String.valueOf(globalcount),new Object[] { roz.getCell(0)+"-P","store_view_code","attribute_set_code","product_type","categories","name","description","meta_title","meta_keyword","meta_description","product_websites","short_description","weight","product_online","tax_class_name","hsn_code","visibility","price","shelf_life","fregrance_color_name","gender","product_contains_liquid","sm2_product_features","dimensions","country","size","unit","ref","brand_name","is_featured","is_bestseller","occasion","scent","manufacturer_detail","packer_detail","manufacturing_name","fragrance_family","fragrence_personality","authenticity","configurable_variations","configurable_variation_labels"});
                      temp=ncs.get(i);
                      globalcount=globalcount+1;
                    }
                    else {
                      temp=ncs.get(i);
                      
                      if(i==0){
                          
                      
                     Row roz=mySheet.getRow(ncs.get(i).getNodenumber());  
                      System.out.println("Node ID:"+temp.getNodenumber()+" Node Classifier:"+temp.getClassifier()+" Value of SKU:"+roz.getCell(0)+"-P"+"  Global Count:"+globalcount);
                      
                      Cell cellsku=rowex.createCell(0);
                      cellsku.setCellValue(roz.getCell(0,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim()+"-P");
                      
                      Cell cellstoreview=rowex.createCell(1);
                      cellstoreview.setCellValue("");
                      
                      Cell cellattribsetcode=rowex.createCell(2);
                      cellattribsetcode.setCellValue(roz.getCell(2,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
                      
                      Cell cellproducttype=rowex.createCell(3);
                      cellproducttype.setCellValue("configurable");
                     
                      Cell cellcategories=rowex.createCell(4);
                      cellcategories.setCellValue("");
                     
                      Cell cellname=rowex.createCell(5);
                      cellname.setCellValue(roz.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
                
                      Cell celldesc=rowex.createCell(6);
                      celldesc.setCellValue(roz.getCell(8,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                      
                      Cell cellmeta=rowex.createCell(7);
                      cellmeta.setCellValue(roz.getCell(15,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
                      
                      Cell cellmetakeywords=rowex.createCell(8);
                      cellmetakeywords.setCellValue(roz.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                      
                      Cell cellmetadesc=rowex.createCell(9);
                      cellmetadesc.setCellValue(roz.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      Cell cellprodweb=rowex.createCell(10);
                      cellprodweb.setCellValue("base");
                      
                      Cell cellshortdesc=rowex.createCell(11);
                      cellshortdesc.setCellValue("");  
                      
                      
                      Cell cellweight=rowex.createCell(12);
                      StringTokenizer stkwt=new StringTokenizer(roz.getCell(10,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString(),".");
                      cellweight.setCellValue(stkwt.nextToken());           
                      

                      Cell cellprodonline=rowex.createCell(13);
                      cellprodonline.setCellValue("1");              

                      Cell celltaxclass=rowex.createCell(14);
                      celltaxclass.setCellValue(roz.getCell(12,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellhsn=rowex.createCell(15);
                      cellhsn.setCellStyle(cellStyle);
                      cellhsn.setCellValue(roz.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
                      
                      Cell cellvisibility=rowex.createCell(16);
                      cellvisibility.setCellValue("Catalog, Search");

                      
                      Cell cellprice=rowex.createCell(17);
                      cellprice.setCellValue(roz.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellshelflife=rowex.createCell(18);
                      cellshelflife.setCellValue(roz.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellfragrancecolor=rowex.createCell(19);
                      cellfragrancecolor.setCellValue("Transparent");
                      
                      
                      Cell cellgender=rowex.createCell(20);
                      cellgender.setCellValue(roz.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                      
                      Cell cellcontainsliquid=rowex.createCell(21);
                      cellcontainsliquid.setCellValue(roz.getCell(23,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      
                      Cell sm2prodfeatures=rowex.createCell(22);
                      sm2prodfeatures.setCellValue(roz.getCell(34,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      Cell celldimensions=rowex.createCell(23);
                      celldimensions.setCellValue(roz.getCell(46,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      
                      Cell cellcountry=rowex.createCell(24);
                      cellcountry.setCellValue(roz.getCell(47,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());   
                      
                      
                      Cell cellbotsize=rowex.createCell(25);
                      String poko=roz.getCell(27,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                    //  StringTokenizer stk=new StringTokenizer(poko,".");
                      cellbotsize.setCellValue(poko);
                      
                      Cell cellunits=rowex.createCell(26);
                      cellunits.setCellValue(roz.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      
                      Cell cellrefs=rowex.createCell(27);
                      StringTokenizer strefs=new StringTokenizer(roz.getCell(53,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString(),".");
                      cellrefs.setCellValue(strefs.nextToken());
                      
                      Cell cellbrands=rowex.createCell(28);
                      cellbrands.setCellValue(roz.getCell(54,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      Cell cellisfeatured=rowex.createCell(29);
                      cellisfeatured.setCellValue((roz.getCell(55,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().equals("Y")?"yes":"no"));

                      
                      Cell cellisbestseller=rowex.createCell(30);
                      cellisbestseller.setCellValue("");
                      
                      Cell celloccasion=rowex.createCell(31);
                      celloccasion.setCellValue(roz.getCell(56,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell cellscent=rowex.createCell(32);
                      cellscent.setCellValue(roz.getCell(57,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      Cell cellmanufacturer=rowex.createCell(33);
                      cellmanufacturer.setCellValue(roz.getCell(58,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell cellpacker=rowex.createCell(34);
                      cellpacker.setCellValue(roz.getCell(59,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell cellmanufname=rowex.createCell(35);
                      cellmanufname.setCellValue(roz.getCell(60,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      Cell cellfragfamily=rowex.createCell(36);
                      cellfragfamily.setCellValue(roz.getCell(20,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      Cell cellfragpersona=rowex.createCell(37);
                      cellfragpersona.setCellValue(roz.getCell(21,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      Cell cellauthenticity=rowex.createCell(38);
                      cellauthenticity.setCellValue(roz.getCell(64,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                    //  Row rowold=writersheet.getRow(globalcount-(childs.size()));
                      
                      
                      String genrated="";
                      int sizex=(childs.size()-1);
                      for(int chx=0;chx<childs.size();chx++) {
                          Row rowtmp=mySheet.getRow(childs.get(chx).getNodenumber());
                          System.out.println("FOUND THE NOS AS FOLLOWS:"+childs.get(chx).getNodenumber());
                          if(sizex==chx){
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",size="+rowtmp.getCell(27).toString()+"";
                          }
                          else {
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",size="+rowtmp.getCell(27).toString()+"|";
                          }
                          System.out.println("Generated String as follows.");
                      }
                      
                     // Cell configurablevariations=rowold.createCell(39);
                     // configurablevariations.setCellValue(genrated);
                      
                      Cell configurablevariationslabel=rowex.createCell(40);
                      configurablevariationslabel.setCellValue("size=Size");
                      
                      lastparentid=globalcount;    
                      childs.clear();
//datawriter(String.valueOf(globalcount),new Object[] { roz.getCell(0)+"-P","store_view_code","attribute_set_code","product_type","categories","name","description","meta_title","meta_keyword","meta_description","product_websites","short_description","weight","product_online","tax_class_name","hsn_code","visibility","price","shelf_life","fregrance_color_name","gender","product_contains_liquid","sm2_product_features","dimensions","country","size","unit","ref","brand_name","is_featured","is_bestseller","occasion","scent","manufacturer_detail","packer_detail","manufacturing_name","fragrance_family","fragrence_personality","authenticity","configurable_variations","configurable_variation_labels"});
                      temp=ncs.get(i);
                      globalcount=globalcount+1;
                          
                          
                          
                      }
                      //childs.clear();
                      //globalcount=globalcount+1;
                    }
//                    System.out.println("Node ID:"+ncs.get(i).getNodenumber()+" Node Classifier:"+ncs.get(i).getClassifier());
                }
                else if(ncs.get(i).getClassifier().equals("Child")) {
                Row rox=mySheet.getRow(ncs.get(i).getNodenumber()); 
                
                childs.add(ncs.get(i));
                
                System.out.println("Node ID:"+ncs.get(i).getNodenumber()+" Node Classifier:"+ncs.get(i).getClassifier()+" Value of SKU:"+rox.getCell(0)+"  Global Count:"+globalcount);
                Cell cellsku=rowex.createCell(0);
                cellsku.setCellValue(rox.getCell(0,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                Cell cellstoreview=rowex.createCell(1);
                cellstoreview.setCellValue("");
                
                Cell cellattribsetcode=rowex.createCell(2);
                cellattribsetcode.setCellValue(rox.getCell(2,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());

                Cell cellproducttype=rowex.createCell(3);
                cellproducttype.setCellValue("simple");
                
                Cell cellcategories=rowex.createCell(4);
                cellcategories.setCellValue("");
                
                
                String poko=rox.getCell(27,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                StringTokenizer stke=new StringTokenizer(poko,".");
                //String sizeam=stke.nextToken();
                String sizeam=poko;
//                cellbotsize.setCellValue(rox.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());


                String mlsizer=rox.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();

                
                
                Cell cellname=rowex.createCell(5);
                cellname.setCellValue(rox.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()+" "+sizeam+" "+mlsizer); 
             
                
                Cell celldesc=rowex.createCell(6);
                celldesc.setCellValue(rox.getCell(8,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());            
            
                Cell cellmeta=rowex.createCell(7);
                cellmeta.setCellValue(rox.getCell(15,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                
                
                Cell cellmetakeywords=rowex.createCell(8);
                cellmetakeywords.setCellValue(rox.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
                Cell cellmetadesc=rowex.createCell(9);
                cellmetadesc.setCellValue(rox.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                
                Cell cellprodweb=rowex.createCell(10);
                cellprodweb.setCellValue("base");
                
                Cell cellshortdesc=rowex.createCell(11);
                cellshortdesc.setCellValue("");  
                
                
                Cell cellweight=rowex.createCell(12);
                
                StringTokenizer stkwt=new StringTokenizer(rox.getCell(10,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString(),".");
                cellweight.setCellValue(stkwt.nextToken());

                Cell cellprodonline=rowex.createCell(13);
                cellprodonline.setCellValue("1");              

                 Cell celltaxclass=rowex.createCell(14);
                 celltaxclass.setCellValue(rox.getCell(12,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                Cell cellhsn=rowex.createCell(15);
                cellhsn.setCellStyle(cellStyle);
                cellhsn.setCellValue(rox.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                 
                
                Cell cellvisibility=rowex.createCell(16);
                cellvisibility.setCellValue("Not Visible Individually");
                
                Cell cellprice=rowex.createCell(17);
                cellprice.setCellValue(rox.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                 
   
                Cell cellshelflife=rowex.createCell(18);
                cellshelflife.setCellValue(rox.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                
                Cell cellfragrancecolor=rowex.createCell(19);
                cellfragrancecolor.setCellValue("Transparent");
                
                Cell cellgender=rowex.createCell(20);
                cellgender.setCellValue(rox.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                
                
                Cell cellcontainsliquid=rowex.createCell(21);
                cellcontainsliquid.setCellValue(rox.getCell(23,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                Cell sm2prodfeatures=rowex.createCell(22);
                sm2prodfeatures.setCellValue(rox.getCell(34,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                
                
                Cell celldimensions=rowex.createCell(23);
                celldimensions.setCellValue(rox.getCell(46,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      
                Cell cellcountry=rowex.createCell(24);
                cellcountry.setCellValue(rox.getCell(47,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                Cell cellbotsize=rowex.createCell(25);
                String pokon=rox.getCell(27,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                StringTokenizer stk=new StringTokenizer(pokon,".");
                cellbotsize.setCellValue(pokon);
//                cellbotsize.setCellValue(rox.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                Cell cellunits=rowex.createCell(26);
                cellunits.setCellValue(rox.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                
                Cell cellrefs=rowex.createCell(27);
                StringTokenizer strefs=new StringTokenizer(rox.getCell(53,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString(),".");
                cellrefs.setCellValue(strefs.nextToken());

            
                Cell cellbrands=rowex.createCell(28);
                cellbrands.setCellValue(rox.getCell(54,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                
                Cell cellisfeatured=rowex.createCell(29);
                cellisfeatured.setCellValue((rox.getCell(55,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().equals("Y")?"yes":"no"));


                Cell cellisbestseller=rowex.createCell(30);
                cellisbestseller.setCellValue("");
                
                Cell celloccasion=rowex.createCell(31);
                celloccasion.setCellValue(rox.getCell(56,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                Cell cellscent=rowex.createCell(32);
                cellscent.setCellValue(rox.getCell(57,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
   
                Cell cellmanufacturer=rowex.createCell(33);
                cellmanufacturer.setCellValue(rox.getCell(58,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
   
                Cell cellpacker=rowex.createCell(34);
                cellpacker.setCellValue(rox.getCell(59,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
        
                Cell cellmanufname=rowex.createCell(35);
                cellmanufname.setCellValue(rox.getCell(60,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                Cell cellfragfamily=rowex.createCell(36);
                cellfragfamily.setCellValue(rox.getCell(20,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                
                Cell cellfragpersona=rowex.createCell(37);
                cellfragpersona.setCellValue(rox.getCell(21,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                
                
                 Cell cellauthenticity=rowex.createCell(38);
                 cellauthenticity.setCellValue(rox.getCell(64,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
        
                
                
                //datawriter.put(String.valueOf(globalcount),new Object[] { rox.getCell(0).toString(),"store_view_code","attribute_set_code","product_type","categories","name","description","meta_title","meta_keyword","meta_description","product_websites","short_description","weight","product_online","tax_class_name","hsn_code","visibility","price","shelf_life","fregrance_color_name","gender","product_contains_liquid","sm2_product_features","dimensions","country","size","unit","ref","brand_name","is_featured","is_bestseller","occasion","scent","manufacturer_detail","packer_detail","manufacturing_name","fragrance_family","fragrence_personality","authenticity","configurable_variations","configurable_variation_labels"});
                globalcount=globalcount+1;
                }
                
            }
            
            
            Row rowexl = writersheet.createRow(globalcount);
            
            Row ron=mySheet.getRow(temp.getNodenumber());
            System.out.println("Node ID:"+temp.getNodenumber()+" Node Classifier:"+temp.getClassifier()+" Value of SKU:"+ron.getCell(0)+"-P"+"  Global Count:"+globalcount);
            //datawriter.put(String.valueOf(globalcount),new Object[] { ron.getCell(0)+"-P","store_view_code","attribute_set_code","product_type","categories","name","description","meta_title","meta_keyword","meta_description","product_websites","short_description","weight","product_online","tax_class_name","hsn_code","visibility","price","shelf_life","fregrance_color_name","gender","product_contains_liquid","sm2_product_features","dimensions","country","size","unit","ref","brand_name","is_featured","is_bestseller","occasion","scent","manufacturer_detail","packer_detail","manufacturing_name","fragrance_family","fragrence_personality","authenticity","configurable_variations","configurable_variation_labels"});
            Cell cellsku=rowexl.createCell(0);
            cellsku.setCellValue(ron.getCell(0,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim()+"-P");

            Cell cellstoreview=rowexl.createCell(1);
            cellstoreview.setCellValue("");

            
            Cell cellattribsetcode=rowexl.createCell(2);
            cellattribsetcode.setCellValue(ron.getCell(2,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

            Cell cellproducttype=rowexl.createCell(3);
            cellproducttype.setCellValue("configurable");

            Cell cellcategories=rowexl.createCell(4);
            cellcategories.setCellValue("");

            Cell cellname=rowexl.createCell(5);
            cellname.setCellValue(ron.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());  
            
            Cell celldesc=rowexl.createCell(6);
            celldesc.setCellValue(ron.getCell(8,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());            
            
            
            Cell cellmeta=rowexl.createCell(7);
            cellmeta.setCellValue(ron.getCell(15,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellmetakeywords=rowexl.createCell(8);
            cellmetakeywords.setCellValue(ron.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellmetadesc=rowexl.createCell(9);
            cellmetadesc.setCellValue(ron.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
            Cell cellprodweb=rowexl.createCell(10);
            cellprodweb.setCellValue("base");  
            
            Cell cellshortdesc=rowexl.createCell(11);
            cellshortdesc.setCellValue("");  

            Cell cellweight=rowexl.createCell(12);

            StringTokenizer stkwt=new StringTokenizer(ron.getCell(10,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString(),".");
            cellweight.setCellValue(stkwt.nextToken());
            

            Cell cellprodonline=rowexl.createCell(13);
            cellprodonline.setCellValue("1");   
            
            Cell celltaxclass=rowexl.createCell(14);
            celltaxclass.setCellValue(ron.getCell(12,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellhsn=rowexl.createCell(15);
            cellhsn.setCellStyle(cellStyle);
            DataFormatter dataFormatter = new DataFormatter();
            
            String formattedCellStr = dataFormatter.formatCellValue(ron.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK));

          //  System.out.println("HSN Code is "+ron.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getLo);
            cellhsn.setCellValue(formattedCellStr);           

            Cell cellvisibility=rowexl.createCell(16);
            cellvisibility.setCellValue("Catalog, Search");
            
            
            Cell cellprice=rowexl.createCell(17);
            cellprice.setCellValue(ron.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellshelflife=rowexl.createCell(18);
            cellshelflife.setCellValue(ron.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellfragrancecolor=rowexl.createCell(19);
            cellfragrancecolor.setCellValue("Transparent");
            
            Cell cellgender=rowexl.createCell(20);
            cellgender.setCellValue(ron.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            
            Cell cellcontainsliquid=rowexl.createCell(21);
            cellcontainsliquid.setCellValue(ron.getCell(23,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            
            Cell sm2prodfeatures=rowexl.createCell(22);
            sm2prodfeatures.setCellValue(ron.getCell(34,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            
            Cell celldimensions=rowexl.createCell(23);
            celldimensions.setCellValue(ron.getCell(46,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            
            Cell cellcountry=rowexl.createCell(24);
            cellcountry.setCellValue(ron.getCell(47,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellbotsize=rowexl.createCell(25);
            String poko=ron.getCell(27,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
            StringTokenizer stk=new StringTokenizer(poko,".");
            cellbotsize.setCellValue(poko);
           // cellbotsize.setCellValue(ron.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
             
            
            Cell cellunits=rowexl.createCell(26);
            cellunits.setCellValue(ron.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
            
            
            Cell cellrefs=rowexl.createCell(27);
            cellrefs.setCellValue(ron.getCell(53,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

            
            Cell cellbrands=rowexl.createCell(28);
            cellbrands.setCellValue(ron.getCell(54,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

            Cell cellisfeatured=rowexl.createCell(29);
            cellisfeatured.setCellValue((ron.getCell(55,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().equals("Y")?"yes":"no"));

            
            Cell cellisbestseller=rowexl.createCell(30);
            cellisbestseller.setCellValue("");
                
            Cell celloccasion=rowexl.createCell(31);
            celloccasion.setCellValue(ron.getCell(56,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

            Cell cellscent=rowexl.createCell(32);
            cellscent.setCellValue(ron.getCell(57,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

            Cell cellmanufacturer=rowexl.createCell(33);
            cellmanufacturer.setCellValue(ron.getCell(58,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
   
            Cell cellpacker=rowexl.createCell(34);
            cellpacker.setCellValue(ron.getCell(59,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                
            Cell cellmanufname=rowexl.createCell(35);
            cellmanufname.setCellValue(ron.getCell(60,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
            
            Cell cellfragfamily=rowexl.createCell(36);
            cellfragfamily.setCellValue(ron.getCell(20,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

            
            Cell cellfragpersona=rowexl.createCell(37);
            cellfragpersona.setCellValue(ron.getCell(21,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
            
            Cell cellauthenticity=rowexl.createCell(38);
            cellauthenticity.setCellValue(ron.getCell(64,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
            
            
            
             String genrated="";
            int sizex=(childs.size()-1);
            for(int chx=0;chx<childs.size();chx++) {
                Row rowtmp=mySheet.getRow(childs.get(chx).getNodenumber());
                if(sizex==chx){
                genrated=genrated+rowtmp.getCell(0).toString()+","+rowtmp.getCell(27).toString()+"";
                }
                else {
                genrated=genrated+rowtmp.getCell(0).toString()+","+rowtmp.getCell(27).toString()+"|";
                }
            }

            Cell configurablevariations=rowexl.createCell(39);
            configurablevariations.setCellValue(genrated);

            Cell configurablevariationslabel=rowexl.createCell(40);
            configurablevariationslabel.setCellValue("size=Size");
            
            
            childs.clear();
// ArrayList<String> keyset = datawriter; 
            
        int rownumx = 0;
 
// 
// 
//        for (String key : keyset) {
//  
//            // Creating a new row in the sheet
//            Row rowx = writersheet.createRow(rownumx++);
//  
//            Object[] objArr = datawriter.get(key);
//  
//            int cellnumx = 0;
//  
//            for (Object obj : objArr) {
//  
//                // This line creates a cell in the next
//                //  column of that row
//                Cell cellx = rowx.createCell(cellnumx++);
//  
//                if (obj instanceof String)
//                    cellx.setCellValue((String)obj);
//  
//                else if (obj instanceof Integer)
//                    cellx.setCellValue((Integer)obj);
//            }
//        }

        


            
            

            
        try (FileOutputStream outputStream = new FileOutputStream("/home/narayan/MakeupDump.xlsx")) {
            writerworkbook.write(outputStream);
        }

            
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ParcosImportSys.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ParcosImportSys.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fis.close();
            } catch (IOException ex) {
                Logger.getLogger(ParcosImportSys.class.getName()).log(Level.SEVERE, null, ex);
            }
        }


    }
    
    
}
