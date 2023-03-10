/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
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

/**
 *
 * @author narayan
 */
public class ParcosImportSys {

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
        Object[][] datawriter = {{ "sku","store_view_code","attribute_set_code","product_type","categories","name","description","meta_title","meta_keyword","meta_description","product_websites","short_description","weight","product_online","tax_class_name","hsn_code","visibility","price","shelf_life","fregrance_color_name","gender","product_contains_liquid","sm2_product_features","dimensions","country","size","unit","ref","brand_name","is_featured","is_bestseller","occasion","scent","manufacturer_detail","packer_detail","manufacturing_name","fragrance_family","fragrence_personality","authenticity","configurable_variations","configurable_variation_labels","base_image_URL (Desktop)","base_image Label","Additional_images (Comma separated)","base_image_mobile_URL","base_image_mobile_label","Listing page Image (small_image) 300 X300","Listing page Image Label (small_image_label)","My Account page Image  (thumbnail_image) 100 X100","My Account page Image  Label (thumbnail_image_label)","Hover_image Listing Page  300X300 (sm2_hover_image)","how_to_user_img1","Desc_image1","Desc_image1_alt1","Desc_image2","Desc_image2_alt2","Desc_image3","Desc_image3_alt3","Desc_image4","Desc_image4_alt4","Desc_image5","Desc_image5_alt5","Desc_image6","Desc_image6_alt6","Desc_image7","Desc_image7_alt7","Desc_image8","Desc_image8_alt8","Desc_image9","Desc_image9_alt9","Desc_image10","Desc_image10_alt10","Desc_image1_mobile","Desc_image1_alt1_mobile","Desc_image2_mobile","Desc_image2_alt2_mobile","Desc_image3_mobile","Desc_image3_alt3_mobile","Desc_image4_mobile","Desc_image4_alt4_mobile","Desc_image5_mobile","Desc_image5_alt5_mobile","Desc_image6_mobile","Desc_image6_alt6_mobile","Desc_image7_mobile","Desc_image7_alt7_mobile","Desc_image8","Desc_image8_alt8_mobile","Desc_image9_mobile","Desc_image9_alt9_mobile","Desc_image10_mobile","Desc_image10_alt10_mobile"}};
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
            
            
            
            File myFile = new File("/home/narayan/Downloads/Parcosham.xlsx");
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
                      cellmeta.setCellValue(roz.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
                      
                      Cell cellmetakeywords=rowex.createCell(8);
                      cellmetakeywords.setCellValue(roz.getCell(18,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                      
                      Cell cellmetadesc=rowex.createCell(9);
                      cellmetadesc.setCellValue(roz.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

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
                      celltaxclass.setCellValue(roz.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellhsn=rowex.createCell(15);
                      cellhsn.setCellStyle(cellStyle);
                      cellhsn.setCellValue(roz.getCell(14,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
                      
                      Cell cellvisibility=rowex.createCell(16);
                      cellvisibility.setCellValue("Catalog, Search");

                      
                      Cell cellprice=rowex.createCell(17);
                      cellprice.setCellValue(roz.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellshelflife=rowex.createCell(18);
                      cellshelflife.setCellValue(roz.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellfragrancecolor=rowex.createCell(19);
                      cellfragrancecolor.setCellValue("Transparent");
                      
                      
                      Cell cellgender=rowex.createCell(20);
                      cellgender.setCellValue(roz.getCell(22,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                      
                      Cell cellcontainsliquid=rowex.createCell(21);
                      cellcontainsliquid.setCellValue(roz.getCell(23,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      
                      Cell sm2prodfeatures=rowex.createCell(22);
                      sm2prodfeatures.setCellValue(roz.getCell(34,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      Cell celldimensions=rowex.createCell(23);
                      celldimensions.setCellValue(roz.getCell(46,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      
                      Cell cellcountry=rowex.createCell(24);
                      cellcountry.setCellValue(roz.getCell(47,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());   
                      
                      
                      Cell cellbotsize=rowex.createCell(25);
                      String poko=roz.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                      StringTokenizer stk=new StringTokenizer(poko,".");
                      cellbotsize.setCellValue(stk.nextToken());
                      
                      
                      Cell cellunits=rowex.createCell(26);
                      cellunits.setCellValue(roz.getCell(52,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      
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
                            String pokox=rowtmp.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                            StringTokenizer stkx=new StringTokenizer(pokox,".");
                            //cellbotsize.setCellValue(stkx.nextToken());
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",size="+stkx.nextToken()+"";
                          }
                          else {
                              
                            String pokox=rowtmp.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                            StringTokenizer stkx=new StringTokenizer(pokox,".");
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",size="+stkx.nextToken()+"|";
                          }
                          System.out.println("Generated String as follows.");
                      }
                      
                      Cell configurablevariations=rowold.createCell(39);
                      configurablevariations.setCellValue(genrated);
                      
                      Cell configurablevariationslabel=rowex.createCell(40);
                      configurablevariationslabel.setCellValue("size=Size");
                      
                      Cell baseimgurl600=rowex.createCell(41);
                      baseimgurl600.setCellValue(roz.getCell(24,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell baseimglabel=rowex.createCell(42);
                      baseimglabel.setCellValue(roz.getCell(25,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell additionalImgs=rowex.createCell(43);
                      additionalImgs.setCellValue(roz.getCell(26,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell baseimgmobileurl=rowex.createCell(44);
                      baseimgmobileurl.setCellValue(roz.getCell(27,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell baseimgmobilelabel=rowex.createCell(45);
                      baseimgmobilelabel.setCellValue(roz.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell listingpagesmallimg=rowex.createCell(46);
                      listingpagesmallimg.setCellValue(roz.getCell(29,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell listingpagesmallimglabel=rowex.createCell(47);
                      listingpagesmallimglabel.setCellValue(roz.getCell(30,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell myaccountpgimg=rowex.createCell(48);
                      myaccountpgimg.setCellValue(roz.getCell(31,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell myaccountpgimglabel=rowex.createCell(49);
                      myaccountpgimglabel.setCellValue(roz.getCell(32,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell hoverimglistingpg=rowex.createCell(50);
                      hoverimglistingpg.setCellValue(roz.getCell(33,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell howtouseimg1=rowex.createCell(51);
                      howtouseimg1.setCellValue(roz.getCell(34,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg1=rowex.createCell(52);
                      descimg1.setCellValue(roz.getCell(66,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg1alt1=rowex.createCell(53);
                      descimg1alt1.setCellValue(roz.getCell(67,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg2=rowex.createCell(54);
                      descimg2.setCellValue(roz.getCell(68,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg2alt2=rowex.createCell(55);
                      descimg2alt2.setCellValue(roz.getCell(69,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg3=rowex.createCell(56);
                      descimg3.setCellValue(roz.getCell(70,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg3alt3=rowex.createCell(57);
                      descimg3alt3.setCellValue(roz.getCell(71,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg4=rowex.createCell(58);
                      descimg4.setCellValue(roz.getCell(72,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg4alt4=rowex.createCell(59);
                      descimg4alt4.setCellValue(roz.getCell(73,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg5=rowex.createCell(60);
                      descimg5.setCellValue(roz.getCell(74,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg5alt5=rowex.createCell(61);
                      descimg5alt5.setCellValue(roz.getCell(75,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg6=rowex.createCell(62);
                      descimg6.setCellValue(roz.getCell(76,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg6alt6=rowex.createCell(63);
                      descimg6alt6.setCellValue(roz.getCell(77,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg7=rowex.createCell(64);
                      descimg7.setCellValue(roz.getCell(78,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg7alt7=rowex.createCell(65);
                      descimg7alt7.setCellValue(roz.getCell(79,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg8=rowex.createCell(66);
                      descimg8.setCellValue(roz.getCell(80,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg8alt8=rowex.createCell(67);
                      descimg8alt8.setCellValue(roz.getCell(81,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg9=rowex.createCell(68);
                      descimg9.setCellValue(roz.getCell(82,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg9alt9=rowex.createCell(69);
                      descimg9alt9.setCellValue(roz.getCell(83,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg10=rowex.createCell(70);
                      descimg10.setCellValue(roz.getCell(84,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg10alt10=rowex.createCell(71);
                      descimg10alt10.setCellValue(roz.getCell(85,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg1mob=rowex.createCell(72);
                      descimg1mob.setCellValue(roz.getCell(86,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg1mobalt=rowex.createCell(73);
                      descimg1mobalt.setCellValue(roz.getCell(87,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg2mob=rowex.createCell(74);
                      descimg2mob.setCellValue(roz.getCell(88,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg2mobalt=rowex.createCell(75);
                      descimg2mobalt.setCellValue(roz.getCell(89,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg3mob=rowex.createCell(76);
                      descimg3mob.setCellValue(roz.getCell(90,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg3mobalt=rowex.createCell(77);
                      descimg3mobalt.setCellValue(roz.getCell(91,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg4mob=rowex.createCell(78);
                      descimg4mob.setCellValue(roz.getCell(92,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg4mobalt=rowex.createCell(79);
                      descimg4mobalt.setCellValue(roz.getCell(93,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg5mob=rowex.createCell(80);
                      descimg5mob.setCellValue(roz.getCell(94,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg5mobalt=rowex.createCell(81);
                      descimg5mobalt.setCellValue(roz.getCell(95,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg6mob=rowex.createCell(82);
                      descimg6mob.setCellValue(roz.getCell(96,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg6mobalt=rowex.createCell(83);
                      descimg6mobalt.setCellValue(roz.getCell(97,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg7mob=rowex.createCell(84);
                      descimg7mob.setCellValue(roz.getCell(98,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg7mobalt=rowex.createCell(85);
                      descimg7mobalt.setCellValue(roz.getCell(99,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg8mob=rowex.createCell(86);
                      descimg8mob.setCellValue(roz.getCell(100,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg8mobalt=rowex.createCell(87);
                      descimg8mobalt.setCellValue(roz.getCell(101,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg9mob=rowex.createCell(88);
                      descimg9mob.setCellValue(roz.getCell(102,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg9mobalt=rowex.createCell(89);
                      descimg9mobalt.setCellValue(roz.getCell(103,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg10mob=rowex.createCell(90);
                      descimg10mob.setCellValue(roz.getCell(104,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg10mobalt=rowex.createCell(91);
                      descimg10mobalt.setCellValue(roz.getCell(105,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
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
                      cellmeta.setCellValue(roz.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
                      
                      Cell cellmetakeywords=rowex.createCell(8);
                      cellmetakeywords.setCellValue(roz.getCell(18,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                      
                      Cell cellmetadesc=rowex.createCell(9);
                      cellmetadesc.setCellValue(roz.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

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
                      celltaxclass.setCellValue(roz.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellhsn=rowex.createCell(15);
                      cellhsn.setCellStyle(cellStyle);
                      cellhsn.setCellValue(roz.getCell(14,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
                      
                      Cell cellvisibility=rowex.createCell(16);
                      cellvisibility.setCellValue("Catalog, Search");

                      
                      Cell cellprice=rowex.createCell(17);
                      cellprice.setCellValue(roz.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellshelflife=rowex.createCell(18);
                      cellshelflife.setCellValue(roz.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                      Cell cellfragrancecolor=rowex.createCell(19);
                      cellfragrancecolor.setCellValue("Transparent");
                      
                      
                      Cell cellgender=rowex.createCell(20);
                      cellgender.setCellValue(roz.getCell(22,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                      
                      Cell cellcontainsliquid=rowex.createCell(21);
                      cellcontainsliquid.setCellValue(roz.getCell(23,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      
                      Cell sm2prodfeatures=rowex.createCell(22);
                      sm2prodfeatures.setCellValue(roz.getCell(34,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      Cell celldimensions=rowex.createCell(23);
                      celldimensions.setCellValue(roz.getCell(46,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      
                      Cell cellcountry=rowex.createCell(24);
                      cellcountry.setCellValue(roz.getCell(47,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());   
                      
                      
                      Cell cellbotsize=rowex.createCell(25);
                      String poko=roz.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                      StringTokenizer stk=new StringTokenizer(poko,".");
                      cellbotsize.setCellValue(stk.nextToken());
                      
                      Cell cellunits=rowex.createCell(26);
                      cellunits.setCellValue(roz.getCell(52,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                      
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
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",size="+rowtmp.getCell(51).toString()+"";
                          }
                          else {
                          genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",size="+rowtmp.getCell(51).toString()+"|";
                          }
                          System.out.println("Generated String as follows.");
                      }
                      
                     // Cell configurablevariations=rowold.createCell(39);
                     // configurablevariations.setCellValue(genrated);
                      
                      Cell configurablevariationslabel=rowex.createCell(40);
                      configurablevariationslabel.setCellValue("size=Size");
                      
                      Cell baseimgurl600=rowex.createCell(41);
                      baseimgurl600.setCellValue(roz.getCell(24,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell baseimglabel=rowex.createCell(42);
                      baseimglabel.setCellValue(roz.getCell(25,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell additionalImgs=rowex.createCell(43);
                      additionalImgs.setCellValue(roz.getCell(26,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell baseimgmobileurl=rowex.createCell(44);
                      baseimgmobileurl.setCellValue(roz.getCell(27,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell baseimgmobilelabel=rowex.createCell(45);
                      baseimgmobilelabel.setCellValue(roz.getCell(28,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell listingpagesmallimg=rowex.createCell(46);
                      listingpagesmallimg.setCellValue(roz.getCell(29,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell listingpagesmallimglabel=rowex.createCell(47);
                      listingpagesmallimglabel.setCellValue(roz.getCell(30,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell myaccountpgimg=rowex.createCell(48);
                      myaccountpgimg.setCellValue(roz.getCell(31,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell myaccountpgimglabel=rowex.createCell(49);
                      myaccountpgimglabel.setCellValue(roz.getCell(32,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell hoverimglistingpg=rowex.createCell(50);
                      hoverimglistingpg.setCellValue(roz.getCell(33,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell howtouseimg1=rowex.createCell(51);
                      howtouseimg1.setCellValue(roz.getCell(34,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg1=rowex.createCell(52);
                      descimg1.setCellValue(roz.getCell(66,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg1alt1=rowex.createCell(53);
                      descimg1alt1.setCellValue(roz.getCell(67,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg2=rowex.createCell(54);
                      descimg2.setCellValue(roz.getCell(68,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg2alt2=rowex.createCell(55);
                      descimg2alt2.setCellValue(roz.getCell(69,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg3=rowex.createCell(56);
                      descimg3.setCellValue(roz.getCell(70,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg3alt3=rowex.createCell(57);
                      descimg3alt3.setCellValue(roz.getCell(71,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg4=rowex.createCell(58);
                      descimg4.setCellValue(roz.getCell(72,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg4alt4=rowex.createCell(59);
                      descimg4alt4.setCellValue(roz.getCell(73,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg5=rowex.createCell(60);
                      descimg5.setCellValue(roz.getCell(74,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg5alt5=rowex.createCell(61);
                      descimg5alt5.setCellValue(roz.getCell(75,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg6=rowex.createCell(62);
                      descimg6.setCellValue(roz.getCell(76,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg6alt6=rowex.createCell(63);
                      descimg6alt6.setCellValue(roz.getCell(77,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg7=rowex.createCell(64);
                      descimg7.setCellValue(roz.getCell(78,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg7alt7=rowex.createCell(65);
                      descimg7alt7.setCellValue(roz.getCell(79,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg8=rowex.createCell(66);
                      descimg8.setCellValue(roz.getCell(80,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg8alt8=rowex.createCell(67);
                      descimg8alt8.setCellValue(roz.getCell(81,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg9=rowex.createCell(68);
                      descimg9.setCellValue(roz.getCell(82,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg9alt9=rowex.createCell(69);
                      descimg9alt9.setCellValue(roz.getCell(83,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg10=rowex.createCell(70);
                      descimg10.setCellValue(roz.getCell(84,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg10alt10=rowex.createCell(71);
                      descimg10alt10.setCellValue(roz.getCell(85,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg1mob=rowex.createCell(72);
                      descimg1mob.setCellValue(roz.getCell(86,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg1mobalt=rowex.createCell(73);
                      descimg1mobalt.setCellValue(roz.getCell(87,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg2mob=rowex.createCell(74);
                      descimg2mob.setCellValue(roz.getCell(88,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg2mobalt=rowex.createCell(75);
                      descimg2mobalt.setCellValue(roz.getCell(89,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg3mob=rowex.createCell(76);
                      descimg3mob.setCellValue(roz.getCell(90,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg3mobalt=rowex.createCell(77);
                      descimg3mobalt.setCellValue(roz.getCell(91,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg4mob=rowex.createCell(78);
                      descimg4mob.setCellValue(roz.getCell(92,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg4mobalt=rowex.createCell(79);
                      descimg4mobalt.setCellValue(roz.getCell(93,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg5mob=rowex.createCell(80);
                      descimg5mob.setCellValue(roz.getCell(94,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg5mobalt=rowex.createCell(81);
                      descimg5mobalt.setCellValue(roz.getCell(95,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg6mob=rowex.createCell(82);
                      descimg6mob.setCellValue(roz.getCell(96,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg6mobalt=rowex.createCell(83);
                      descimg6mobalt.setCellValue(roz.getCell(97,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg7mob=rowex.createCell(84);
                      descimg7mob.setCellValue(roz.getCell(98,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg7mobalt=rowex.createCell(85);
                      descimg7mobalt.setCellValue(roz.getCell(99,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg8mob=rowex.createCell(86);
                      descimg8mob.setCellValue(roz.getCell(100,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg8mobalt=rowex.createCell(87);
                      descimg8mobalt.setCellValue(roz.getCell(101,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg9mob=rowex.createCell(88);
                      descimg9mob.setCellValue(roz.getCell(102,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg9mobalt=rowex.createCell(89);
                      descimg9mobalt.setCellValue(roz.getCell(103,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg10mob=rowex.createCell(90);
                      descimg10mob.setCellValue(roz.getCell(104,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
                      Cell descimg10mobalt=rowex.createCell(91);
                      descimg10mobalt.setCellValue(roz.getCell(105,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                      
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
                
                
                String poko=rox.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                StringTokenizer stke=new StringTokenizer(poko,".");
                String sizeam=stke.nextToken();
//                cellbotsize.setCellValue(rox.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());


                String mlsizer=rox.getCell(52,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();

                
                
                Cell cellname=rowex.createCell(5);
                cellname.setCellValue(rox.getCell(7,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()+" "+sizeam+" "+mlsizer); 
             
                
                Cell celldesc=rowex.createCell(6);
                celldesc.setCellValue(rox.getCell(8,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());            
            
                Cell cellmeta=rowex.createCell(7);
                cellmeta.setCellValue(rox.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                
                
                Cell cellmetakeywords=rowex.createCell(8);
                cellmetakeywords.setCellValue(rox.getCell(18,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
                Cell cellmetadesc=rowex.createCell(9);
                cellmetadesc.setCellValue(rox.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                
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
                 celltaxclass.setCellValue(rox.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                Cell cellhsn=rowex.createCell(15);
                cellhsn.setCellStyle(cellStyle);
                cellhsn.setCellValue(rox.getCell(14,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                 
                
                Cell cellvisibility=rowex.createCell(16);
                cellvisibility.setCellValue("Not Visible Individually");
                
                Cell cellprice=rowex.createCell(17);
                cellprice.setCellValue(rox.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                 
   
                Cell cellshelflife=rowex.createCell(18);
                cellshelflife.setCellValue(rox.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                
                Cell cellfragrancecolor=rowex.createCell(19);
                cellfragrancecolor.setCellValue("Transparent");
                
                Cell cellgender=rowex.createCell(20);
                cellgender.setCellValue(rox.getCell(22,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
                
                
                Cell cellcontainsliquid=rowex.createCell(21);
                cellcontainsliquid.setCellValue(rox.getCell(23,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                Cell sm2prodfeatures=rowex.createCell(22);
                sm2prodfeatures.setCellValue(rox.getCell(34,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                
                
                Cell celldimensions=rowex.createCell(23);
                celldimensions.setCellValue(rox.getCell(46,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

                      
                Cell cellcountry=rowex.createCell(24);
                cellcountry.setCellValue(rox.getCell(47,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           


                Cell cellbotsize=rowex.createCell(25);
                String pokon=rox.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                StringTokenizer stk=new StringTokenizer(pokon,".");
                cellbotsize.setCellValue(stk.nextToken());
//                cellbotsize.setCellValue(rox.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                Cell cellunits=rowex.createCell(26);
                cellunits.setCellValue(rox.getCell(52,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                
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
            cellmeta.setCellValue(ron.getCell(17,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellmetakeywords=rowexl.createCell(8);
            cellmetakeywords.setCellValue(ron.getCell(18,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellmetadesc=rowexl.createCell(9);
            cellmetadesc.setCellValue(ron.getCell(19,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           
            
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
            celltaxclass.setCellValue(ron.getCell(13,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellhsn=rowexl.createCell(15);
            cellhsn.setCellStyle(cellStyle);
            cellhsn.setCellValue(ron.getCell(14,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellvisibility=rowexl.createCell(16);
            cellvisibility.setCellValue("Catalog, Search");
            
            
            Cell cellprice=rowexl.createCell(17);
            cellprice.setCellValue(ron.getCell(11,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellshelflife=rowexl.createCell(18);
            cellshelflife.setCellValue(ron.getCell(16,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellfragrancecolor=rowexl.createCell(19);
            cellfragrancecolor.setCellValue("Transparent");
            
            Cell cellgender=rowexl.createCell(20);
            cellgender.setCellValue(ron.getCell(22,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            
            Cell cellcontainsliquid=rowexl.createCell(21);
            cellcontainsliquid.setCellValue(ron.getCell(23,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            
            Cell sm2prodfeatures=rowexl.createCell(22);
            sm2prodfeatures.setCellValue(ron.getCell(34,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            
            Cell celldimensions=rowexl.createCell(23);
            celldimensions.setCellValue(ron.getCell(46,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            
            Cell cellcountry=rowexl.createCell(24);
            cellcountry.setCellValue(ron.getCell(47,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());           

            Cell cellbotsize=rowexl.createCell(25);
            String poko=ron.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
            StringTokenizer stk=new StringTokenizer(poko,".");
            cellbotsize.setCellValue(stk.nextToken());
           // cellbotsize.setCellValue(ron.getCell(51,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
             
            
            Cell cellunits=rowexl.createCell(26);
            cellunits.setCellValue(ron.getCell(52,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
            
            
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
                genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",size="+rowtmp.getCell(51).toString()+"";
                }
                else {
                genrated=genrated+"sku="+rowtmp.getCell(0).toString()+",size="+rowtmp.getCell(51).toString()+"|";
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

        


            
            

            
        try (FileOutputStream outputStream = new FileOutputStream("/home/narayan/ParcosDumpNew.xlsx")) {
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
