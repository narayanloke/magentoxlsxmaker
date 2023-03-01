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
import java.util.StringTokenizer;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author narayan
 */
public class SkinCareImportSys {
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
        Object[][] datawriter = {{ "sku","store_view_code","attribute_set_code","product_type","categories","product_websites","name","description","short_description","weight","product_online","tax_class_name","visibility","price","special_price","special_price_from_date","special_price_to_date","url_key","meta_title","meta_keywords","meta_description","base_image","base_image_label","small_image","small_image_label","thumbnail_image","thumbnail_image_label","swatch_image","swatch_image_label","created_at","updated_at","new_from_date","new_to_date","display_product_options_in","map_price","msrp_price","map_enabled","gift_message_available","custom_design","custom_design_from","custom_design_to","custom_layout_update","page_layout","product_options_container","msrp_display_actual_price_type","country_of_manufacture","additional_attributes","qty","out_of_stock_qty","use_config_min_qty","is_qty_decimal","allow_backorders","use_config_backorders","min_cart_qty","use_config_min_sale_qty","max_cart_qty","use_config_max_sale_qty","is_in_stock","notify_on_stock_below","use_config_notify_stock_qty","manage_stock","use_config_manage_stock","use_config_qty_increments","qty_increments","use_config_enable_qty_inc","enable_qty_increments","is_decimal_divided","website_id","related_skus","related_position","crosssell_skus","crosssell_position","upsell_skus","upsell_position","additional_images","additional_image_labels","hide_from_product_page","custom_options","bundle_price_type","bundle_sku_type","bundle_price_view","bundle_weight_type","bundle_values","bundle_shipment_type","associated_skus","downloadable_links","downloadable_samples","configurable_variations","configurable_variation_labels","sm2_hover_image" }};
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
            
            
            
            File myFile = new File("/home/narayan/Downloads/skincare2x.xlsx");
            fis = new FileInputStream(myFile);
            // Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
            // Return first sheet from the XLSX workbook
            XSSFSheet mySheet = myWorkBook.getSheetAt(1);
            
            
            int numrows=mySheet.getPhysicalNumberOfRows();
            
            for(row=0;row<numrows;row++) {
                System.out.println(row);
                
              Row currentRow=mySheet.getRow(row);
              
//              if(currentRow==null){
//                  continue;
//              
//              }
              
                int numcols=currentRow.getPhysicalNumberOfCells();
                //parent=false;
              System.out.println("Next Lane Cols"+numcols);
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
                            System.out.println("Added to Single");
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

                      Cell cellcategory=rowex.createCell(4);
                      cellcategory.setCellValue("");
                      
                      Cell cellproductwebsites=rowex.createCell(5);
                      cellproductwebsites.setCellValue("base");
                      
                      Cell cellname=rowex.createCell(6);
                      cellname.setCellValue(roz.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                     
                      Cell celldescription=rowex.createCell(7);
                      celldescription.setCellValue(roz.getCell(8,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
                      
                      Cell cellshortdescription=rowex.createCell(8);
                      cellshortdescription.setCellValue(roz.getCell(9,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
            
                      Cell cellweight=rowex.createCell(9);
                      cellweight.setCellValue(roz.getCell(10).toString());
                      
                      
                      Cell cellproductonline=rowex.createCell(10);
                      cellproductonline.setCellValue("1");
                      
                      Cell celltaxclass=rowex.createCell(11);
                      celltaxclass.setCellValue(roz.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
                      
                      Cell cellvisibility=rowex.createCell(12);
                      cellvisibility.setCellValue("Catalog, Search");
                      
                      Cell cellprice=rowex.createCell(13);
                      cellprice.setCellValue(roz.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell specialprice=rowex.createCell(14);
                      specialprice.setCellValue("");
                      
                      Cell specialpricefromdt=rowex.createCell(15);
                      specialpricefromdt.setCellValue("");
                      
                      Cell specialpricetodt=rowex.createCell(16);
                      specialpricetodt.setCellValue("");
                      
                      
                      Cell cellurlkey=rowex.createCell(17);
                      cellurlkey.setCellValue("");
                      
                      
                      Cell cellmetatitle=rowex.createCell(18);
                      cellmetatitle.setCellValue(roz.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      
                      Cell cellmetakeyword=rowex.createCell(19);
                      cellmetakeyword.setCellValue(roz.getCell(18,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      
                      Cell cellmetadesc=rowex.createCell(20);
                      cellmetadesc.setCellValue(roz.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      
                      Cell cellbaseimg=rowex.createCell(21);
                      cellbaseimg.setCellValue(roz.getCell(56,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      
                      Cell cellbaseimglabel=rowex.createCell(22);
                      cellbaseimglabel.setCellValue(roz.getCell(57,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      
                      Cell cellsmallimage=rowex.createCell(23);
                      cellsmallimage.setCellValue(roz.getCell(61,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());  
                      
                      
                      Cell cellsmallimagelabel=rowex.createCell(24);
                      cellsmallimagelabel.setCellValue(roz.getCell(62,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 

                      
                      Cell cellthumbnail=rowex.createCell(25);
                      cellthumbnail.setCellValue(roz.getCell(63,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 

                      
                      Cell cellthumbnaillabel=rowex.createCell(26);
                      cellthumbnaillabel.setCellValue(roz.getCell(64,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());                       

                      Cell cellswatchimage=rowex.createCell(27);
                      cellswatchimage.setCellValue("");    
                      
                      Cell cellswatchimagelabel=rowex.createCell(28);
                      cellswatchimagelabel.setCellValue("");   
                      
                      Cell cellcountrymanufacturer=rowex.createCell(45);
                      cellcountrymanufacturer.setCellValue(roz.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());   
                      
                      
                      Cell celladdimages=rowex.createCell(74);
                      celladdimages.setCellValue(roz.getCell(54,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());   
                      
                      
                      //Cell celladditionalattribs=rowex.createCell(46);
                      //String rangers="age_range="+roz.getCell(50,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()+",authenticity="+roz.getCell(50,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                      //celladditionalattribs.setCellValue(rangers);
                      
                      
                      Row rowold=writersheet.getRow(lastparentid);
                      
                      System.out.println("Old Row number:"+(i-childs.size()));
                      
                      String genrated="";
                      int sizex=(childs.size()-1);
                      for(int chx=0;chx<childs.size();chx++) {
                          Row rowtmp=mySheet.getRow(childs.get(chx).getNodenumber());
                          System.out.println("FOUND THE NOS AS FOLLOWS:"+childs.get(chx).getNodenumber());
                          if(sizex==chx){
                            String pokox=rowtmp.getCell(25,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                            StringTokenizer stkx=new StringTokenizer(pokox,".");
                            //cellbotsize.setCellValue(stkx.nextToken());
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",shade="+pokox+"";
                          }
                          else {
                              
                            String pokox=rowtmp.getCell(25,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                            StringTokenizer stkx=new StringTokenizer(pokox,".");
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",shade="+pokox+"|";
                          }
                          System.out.println("Generated String as follows.");
                      }
                      
                      Cell configurablevariations=rowold.createCell(87);
                      configurablevariations.setCellValue(genrated);
                      
                      Cell configurablevariationslabel=rowex.createCell(88);
                      configurablevariationslabel.setCellValue("shade=Shade");
                      
                      Cell cellsm2hoverimage=rowex.createCell(89);
                      cellsm2hoverimage.setCellValue(roz.getCell(61,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
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
                     
                      Cell cellcategory=rowex.createCell(4);
                      cellcategory.setCellValue("");
                      
                      Cell cellproductwebsites=rowex.createCell(5);
                      cellproductwebsites.setCellValue("base");
                      
                      Cell cellname=rowex.createCell(6);
                      cellname.setCellValue(roz.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell celldescription=rowex.createCell(7);
                      celldescription.setCellValue(roz.getCell(8,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());

                      Cell cellshortdescription=rowex.createCell(8);
                      cellshortdescription.setCellValue(roz.getCell(9,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());

                      Cell cellweight=rowex.createCell(9);
                      cellweight.setCellValue(roz.getCell(10).toString());

                      Cell cellproductonline=rowex.createCell(10);
                      cellproductonline.setCellValue("1");

                      Cell celltaxclass=rowex.createCell(11);
                      celltaxclass.setCellValue(roz.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
                      
                      Cell cellvisibility=rowex.createCell(12);
                      cellvisibility.setCellValue("Catalog, Search");
                      
                      Cell cellprice=rowex.createCell(13);
                      cellprice.setCellValue(roz.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      
                      Cell specialprice=rowex.createCell(14);
                      specialprice.setCellValue("");
                      
                      Cell specialpricefromdt=rowex.createCell(15);
                      specialpricefromdt.setCellValue("");
                      
                      Cell specialpricetodt=rowex.createCell(16);
                      specialpricetodt.setCellValue("");
                      
                      Cell cellurlkey=rowex.createCell(17);
                      cellurlkey.setCellValue("");
                      
                      
                      
                      Cell cellmetatitle=rowex.createCell(18);
                      cellmetatitle.setCellValue(roz.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      

                      Cell cellmetakeyword=rowex.createCell(19);
                      cellmetakeyword.setCellValue(roz.getCell(18,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      
                      Cell cellmetadesc=rowex.createCell(20);
                      cellmetadesc.setCellValue(roz.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      
                       
                      Cell cellbaseimg=rowex.createCell(21);
                      cellbaseimg.setCellValue(roz.getCell(56,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      
                      Cell cellbaseimglabel=rowex.createCell(22);
                      cellbaseimglabel.setCellValue(roz.getCell(57,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      
                      Cell cellsmallimage=rowex.createCell(23);
                      cellsmallimage.setCellValue(roz.getCell(61,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());   
                      
                      
                      Cell cellsmallimagelabel=rowex.createCell(24);
                      cellsmallimagelabel.setCellValue(roz.getCell(62,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
                      
                      
                      Cell cellthumbnail=rowex.createCell(25);
                      cellthumbnail.setCellValue(roz.getCell(63,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 

                      
                      Cell cellthumbnaillabel=rowex.createCell(26);
                      cellthumbnaillabel.setCellValue(roz.getCell(64,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
                      
                      Cell cellswatchimage=rowex.createCell(27);
                      cellswatchimage.setCellValue("");    
                      
                      Cell cellswatchimagelabel=rowex.createCell(28);
                      cellswatchimagelabel.setCellValue("");   
                      
                      Cell cellcountrymanufacturer=rowex.createCell(45);
                      cellcountrymanufacturer.setCellValue(roz.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
                      
                      Cell celladdimages=rowex.createCell(74);
                      celladdimages.setCellValue(roz.getCell(54,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());  
                      
                      
                      
                      String genrated="";
                      int sizex=(childs.size()-1);
                      for(int chx=0;chx<childs.size();chx++) {
                          Row rowtmp=mySheet.getRow(childs.get(chx).getNodenumber());
                          System.out.println("FOUND THE NOS AS FOLLOWS:"+childs.get(chx).getNodenumber());
                          if(sizex==chx){
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",shade="+rowtmp.getCell(25).toString()+"";
                          }
                          else {
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",shade="+rowtmp.getCell(25).toString()+"|";
                          }
                          System.out.println("Generated String as follows.");
                      }
                      
                     // Cell configurablevariations=rowold.createCell(39);
                     // configurablevariations.setCellValue(genrated);
                      
                      Cell configurablevariationslabel=rowex.createCell(88);
                      configurablevariationslabel.setCellValue("shade=Shade");
                      
                      Cell cellsm2hoverimage=rowex.createCell(89);
                      cellsm2hoverimage.setCellValue(roz.getCell(61,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      
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
                
                else if(ncs.get(i).getClassifier().equals("Single")) {
                    
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
                
                Cell cellcategory=rowex.createCell(4);
                cellcategory.setCellValue("");
                
                Cell cellproductwebsites=rowex.createCell(5);
                cellproductwebsites.setCellValue("base");
                      
                Cell cellname=rowex.createCell(6);
                String botlesize=rox.getCell(27,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();  
                StringTokenizer stkbotlesize=new StringTokenizer(botlesize,".");
                String strcellnameandsize=rox.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()+"  "+stkbotlesize.nextToken()+" "+rox.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                
                cellname.setCellValue(rox.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());        
                
                Cell celldescription=rowex.createCell(7);
                celldescription.setCellValue(rox.getCell(8,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
  
                
                Cell cellshortdescription=rowex.createCell(8);
                cellshortdescription.setCellValue(rox.getCell(9,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());

                
                Cell cellweight=rowex.createCell(9);
                cellweight.setCellValue(rox.getCell(10).toString());
                
                Cell cellproductonline=rowex.createCell(10);
                cellproductonline.setCellValue("1");
                
                
                Cell celltaxclass=rowex.createCell(11);
                celltaxclass.setCellValue(rox.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
                
                
                Cell cellvisibility=rowex.createCell(12);
                cellvisibility.setCellValue("Catalog, Search");
                
                
                Cell cellprice=rowex.createCell(13);
                cellprice.setCellValue(rox.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                
                Cell specialprice=rowex.createCell(14);
                specialprice.setCellValue("");

                Cell specialpricefromdt=rowex.createCell(15);
                specialpricefromdt.setCellValue("");

                Cell specialpricetodt=rowex.createCell(16);
                specialpricetodt.setCellValue("");
                
                
                Cell cellurlkey=rowex.createCell(17);
                cellurlkey.setCellValue("");
                
                
                Cell cellmetatitle=rowex.createCell(18);
                cellmetatitle.setCellValue(rox.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                Cell cellmetakeyword=rowex.createCell(19);
                cellmetakeyword.setCellValue(rox.getCell(18,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                Cell cellmetadesc=rowex.createCell(20);
                cellmetadesc.setCellValue(rox.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                 
                Cell cellbaseimg=rowex.createCell(21);
                cellbaseimg.setCellValue(rox.getCell(56,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                
                Cell cellbaseimglabel=rowex.createCell(22);
                cellbaseimglabel.setCellValue(rox.getCell(57,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                
                Cell cellsmallimage=rowex.createCell(23);
                cellsmallimage.setCellValue(rox.getCell(61,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());    
                
                
                Cell cellsmallimagelabel=rowex.createCell(24);
                cellsmallimagelabel.setCellValue(rox.getCell(62,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
                
                
                Cell cellthumbnail=rowex.createCell(25);
                cellthumbnail.setCellValue(rox.getCell(63,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 

                      
                Cell cellthumbnaillabel=rowex.createCell(26);
                cellthumbnaillabel.setCellValue(rox.getCell(64,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
                      
                
                Cell cellswatchimage=rowex.createCell(27);
                cellswatchimage.setCellValue("");    
                   
                
                Cell cellswatchimagelabel=rowex.createCell(28);
                cellswatchimagelabel.setCellValue("");   

                
                Cell cellcountrymanufacturer=rowex.createCell(45);
                cellcountrymanufacturer.setCellValue(rox.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
                  
                Cell celladdimages=rowex.createCell(74);
                celladdimages.setCellValue(rox.getCell(54,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());                
                
                
                Cell cellsm2hoverimage=rowex.createCell(89);
                cellsm2hoverimage.setCellValue(rox.getCell(61,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                
                //datawriter.put(String.valueOf(globalcount),new Object[] { rox.getCell(0).toString(),"store_view_code","attribute_set_code","product_type","categories","name","description","meta_title","meta_keyword","meta_description","product_websites","short_description","weight","product_online","tax_class_name","hsn_code","visibility","price","shelf_life","fregrance_color_name","gender","product_contains_liquid","sm2_product_features","dimensions","country","size","unit","ref","brand_name","is_featured","is_bestseller","occasion","scent","manufacturer_detail","packer_detail","manufacturing_name","fragrance_family","fragrence_personality","authenticity","configurable_variations","configurable_variation_labels"});
                globalcount=globalcount+1;
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
                
                Cell cellcategory=rowex.createCell(4);
                cellcategory.setCellValue("");
                
                Cell cellproductwebsites=rowex.createCell(5);
                cellproductwebsites.setCellValue("base");
                      
                Cell cellname=rowex.createCell(6);
                String botlesize=rox.getCell(27,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();  
                StringTokenizer stkbotlesize=new StringTokenizer(botlesize,".");
                String strcellnameandsize=rox.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()+"  "+stkbotlesize.nextToken()+" "+rox.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                
                cellname.setCellValue(rox.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());        
                
                Cell celldescription=rowex.createCell(7);
                celldescription.setCellValue(rox.getCell(8,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
  
                
                Cell cellshortdescription=rowex.createCell(8);
                cellshortdescription.setCellValue(rox.getCell(9,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());

                
                Cell cellweight=rowex.createCell(9);
                cellweight.setCellValue(rox.getCell(10).toString());
                
                Cell cellproductonline=rowex.createCell(10);
                cellproductonline.setCellValue("1");
                
                
                Cell celltaxclass=rowex.createCell(11);
                celltaxclass.setCellValue(rox.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
                
                
                Cell cellvisibility=rowex.createCell(12);
                cellvisibility.setCellValue("Catalog, Search");
                
                
                Cell cellprice=rowex.createCell(13);
                cellprice.setCellValue(rox.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                
                Cell specialprice=rowex.createCell(14);
                specialprice.setCellValue("");

                Cell specialpricefromdt=rowex.createCell(15);
                specialpricefromdt.setCellValue("");

                Cell specialpricetodt=rowex.createCell(16);
                specialpricetodt.setCellValue("");
                
                
                Cell cellurlkey=rowex.createCell(17);
                cellurlkey.setCellValue("");
                
                
                Cell cellmetatitle=rowex.createCell(18);
                cellmetatitle.setCellValue(rox.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                Cell cellmetakeyword=rowex.createCell(19);
                cellmetakeyword.setCellValue(rox.getCell(18,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                Cell cellmetadesc=rowex.createCell(20);
                cellmetadesc.setCellValue(rox.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                 
                Cell cellbaseimg=rowex.createCell(21);
                cellbaseimg.setCellValue(rox.getCell(56,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                
                Cell cellbaseimglabel=rowex.createCell(22);
                cellbaseimglabel.setCellValue(rox.getCell(57,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                
                Cell cellsmallimage=rowex.createCell(23);
                cellsmallimage.setCellValue(rox.getCell(61,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());    
                
                
                Cell cellsmallimagelabel=rowex.createCell(24);
                cellsmallimagelabel.setCellValue(rox.getCell(62,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
                
                
                Cell cellthumbnail=rowex.createCell(25);
                cellthumbnail.setCellValue(rox.getCell(63,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 

                      
                Cell cellthumbnaillabel=rowex.createCell(26);
                cellthumbnaillabel.setCellValue(rox.getCell(64,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
                      
                
                Cell cellswatchimage=rowex.createCell(27);
                cellswatchimage.setCellValue("");    
                   
                
                Cell cellswatchimagelabel=rowex.createCell(28);
                cellswatchimagelabel.setCellValue("");   

                
                Cell cellcountrymanufacturer=rowex.createCell(45);
                cellcountrymanufacturer.setCellValue(rox.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
                  
                Cell celladdimages=rowex.createCell(74);
                celladdimages.setCellValue(rox.getCell(54,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());                
                
                
                Cell cellsm2hoverimage=rowex.createCell(89);
                cellsm2hoverimage.setCellValue(rox.getCell(61,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                
                
                //datawriter.put(String.valueOf(globalcount),new Object[] { rox.getCell(0).toString(),"store_view_code","attribute_set_code","product_type","categories","name","description","meta_title","meta_keyword","meta_description","product_websites","short_description","weight","product_online","tax_class_name","hsn_code","visibility","price","shelf_life","fregrance_color_name","gender","product_contains_liquid","sm2_product_features","dimensions","country","size","unit","ref","brand_name","is_featured","is_bestseller","occasion","scent","manufacturer_detail","packer_detail","manufacturing_name","fragrance_family","fragrence_personality","authenticity","configurable_variations","configurable_variation_labels"});
                globalcount=globalcount+1;
                }
                
            }
            
            
//            Row rowexl = writersheet.createRow(globalcount);
//            
//            Row ron=mySheet.getRow(temp.getNodenumber());
//            System.out.println("Node ID:"+temp.getNodenumber()+" Node Classifier:"+temp.getClassifier()+" Value of SKU:"+ron.getCell(0)+"-P"+"  Global Count:"+globalcount);
//            //datawriter.put(String.valueOf(globalcount),new Object[] { ron.getCell(0)+"-P","store_view_code","attribute_set_code","product_type","categories","name","description","meta_title","meta_keyword","meta_description","product_websites","short_description","weight","product_online","tax_class_name","hsn_code","visibility","price","shelf_life","fregrance_color_name","gender","product_contains_liquid","sm2_product_features","dimensions","country","size","unit","ref","brand_name","is_featured","is_bestseller","occasion","scent","manufacturer_detail","packer_detail","manufacturing_name","fragrance_family","fragrence_personality","authenticity","configurable_variations","configurable_variation_labels"});
//            Cell cellsku=rowexl.createCell(0);
//            cellsku.setCellValue(ron.getCell(0,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim()+"-P");
//
//            Cell cellstoreview=rowexl.createCell(1);
//            cellstoreview.setCellValue("");
//
//            
//            Cell cellattribsetcode=rowexl.createCell(2);
//            cellattribsetcode.setCellValue(ron.getCell(2,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
//
//            Cell cellproducttype=rowexl.createCell(3);
//            cellproducttype.setCellValue("configurable");
//
//            Cell cellcategory=rowexl.createCell(4);
//            cellcategory.setCellValue("");
//                      
//            Cell cellproductwebsites=rowexl.createCell(5);
//            cellproductwebsites.setCellValue("base");
//                      
//            Cell cellname=rowexl.createCell(6);
//            cellname.setCellValue(ron.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
//            
//            Cell celldescription=rowexl.createCell(7);
//            celldescription.setCellValue(ron.getCell(8,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
//                      
//            Cell cellshortdescription=rowexl.createCell(8);
//            cellshortdescription.setCellValue(ron.getCell(9,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
//
//            
//            Cell cellweight=rowexl.createCell(9);
//            cellweight.setCellValue(ron.getCell(10,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
//            
//            Cell cellproductonline=rowexl.createCell(10);
//            cellproductonline.setCellValue("1");
//            
//            Cell celltaxclass=rowexl.createCell(11);
//            celltaxclass.setCellValue(ron.getCell(12,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim());
//                      
//            Cell cellvisibility=rowexl.createCell(12);
//            cellvisibility.setCellValue("Catalog,Search");
//            
//            Cell cellprice=rowexl.createCell(13);
//            cellprice.setCellValue(ron.getCell(11).toString());
//            
//            Cell specialprice=rowexl.createCell(14);
//            specialprice.setCellValue("");
//
//            Cell specialpricefromdt=rowexl.createCell(15);
//            specialpricefromdt.setCellValue("");
//
//            Cell specialpricetodt=rowexl.createCell(16);
//            specialpricetodt.setCellValue("");
//            
//            Cell cellurlkey=rowexl.createCell(17);
//            cellurlkey.setCellValue("");
//            
//            Cell cellmetatitle=rowexl.createCell(18);
//            cellmetatitle.setCellValue(ron.getCell(15,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());            
//            
//            Cell cellmetakeyword=rowexl.createCell(19);
//            cellmetakeyword.setCellValue(ron.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
//                                 
//            Cell cellmetadesc=rowexl.createCell(20);
//            cellmetadesc.setCellValue(ron.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
//                                   
//            Cell cellbaseimg=rowexl.createCell(21);
//            cellbaseimg.setCellValue(ron.getCell(52,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
//                      
//            
//            Cell cellbaseimglabel=rowexl.createCell(22);
//            cellbaseimglabel.setCellValue(ron.getCell(53,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
//                      
//            
//            Cell cellsmallimage=rowexl.createCell(22);
//            cellsmallimage.setCellValue(ron.getCell(57,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());    
//            
//            Cell cellsmallimagelabel=rowexl.createCell(24);
//            cellsmallimagelabel.setCellValue(ron.getCell(58,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
//            
//            
//            Cell cellthumbnail=rowexl.createCell(25);
//            cellthumbnail.setCellValue(ron.getCell(59,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
//
//                      
//            Cell cellthumbnaillabel=rowexl.createCell(26);
//            cellthumbnaillabel.setCellValue(ron.getCell(60,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
//            
//            
//            Cell cellswatchimage=rowexl.createCell(27);
//            cellswatchimage.setCellValue("");    
//
//            Cell cellswatchimagelabel=rowexl.createCell(28);
//            cellswatchimagelabel.setCellValue("");   
//            
//            
//            Cell cellcountrymanufacturer=rowexl.createCell(45);
//            cellcountrymanufacturer.setCellValue(ron.getCell(21,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()); 
//                      
//            
//            
//            Cell celladdimages=rowexl.createCell(74);
//            celladdimages.setCellValue(ron.getCell(54,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());            
//                      
//            
//             String genrated="";
//            int sizex=(childs.size()-1);
//            for(int chx=0;chx<childs.size();chx++) {
//                Row rowtmp=mySheet.getRow(childs.get(chx).getNodenumber());
//                if(sizex==chx){
//                genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",shade="+rowtmp.getCell(25).toString()+"";
//                }
//                else {
//                genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",shade="+rowtmp.getCell(25).toString()+"|";
//                }
//            }
//
//            Cell configurablevariations=rowexl.createCell(87);
//            configurablevariations.setCellValue(genrated);
//
//            Cell configurablevariationslabel=rowexl.createCell(88);
//            configurablevariationslabel.setCellValue("shade=Shade");
//            
//            
//            Cell cellsm2hoverimage=rowexl.createCell(89);
//            cellsm2hoverimage.setCellValue(ron.getCell(61,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
//            
//            
//            childs.clear();
// ArrayList<String> keyset = datawriter; 
            
        int rownumx = 0;
     


            
            

            
        try (FileOutputStream outputStream = new FileOutputStream("/home/narayan/SkincareDumper.xlsx")) {
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
