package poiexample;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TableExample {
	
	public static void main(String[] args) throws Exception
	{
		try {

		    FileInputStream file = new FileInputStream(new File("C:/Users/tarun.samanta/OneDrive - Accenture/Desktop/Primavera.xlsx"));
		    Workbook workbook = new XSSFWorkbook(file);
		    Sheet datatypeSheet = workbook.getSheetAt(0);
		    Iterator<Row> rowiterator = datatypeSheet.iterator();
		     String key;
		     String value;
		     while (rowiterator.hasNext()) {
		       Row nextRow = rowiterator.next();
		       Iterator<Cell> cellIterator = nextRow.cellIterator();
		       key= null;
			   value= null;
		       if(nextRow.getCell(0)!=null)
		    	   key = nextRow.getCell(0).getStringCellValue();
		       
		       if(nextRow.getCell(3)!=null)
		    	   value = nextRow.getCell(3).getStringCellValue();
		       
		       if(key==null)
		    	   continue;
		      
		       else if(value!=null && value.equals("Removed"))
		    	   continue;
		       else if(key!=null && key.length()>5 && key.substring(0,4).equals("http")) {
		    	   
		    	   URL url=new URL("https://www.linkedin.com/");
		    	   HttpURLConnection connection=(HttpURLConnection)url.openConnection();
		    	   connection.setRequestMethod("GET");
		    	   connection.connect();
		    	   int code=connection.getResponseCode();
		    	   
		    	   if(code==200) {
		    		   nextRow.getCell(3).setCellValue("OK");
		    	   }
		       }  
		     }
		     file.close();
		     FileOutputStream outFile = new FileOutputStream(new File("C:/Users/tarun.samanta/OneDrive - Accenture/Desktop/Primavera.xlsx"));
		     workbook.write(outFile);
		     outFile.close();
		     workbook.close();
		     
		     System.out.println("Success");

            
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
		
		
		
	}

		   
	
}
