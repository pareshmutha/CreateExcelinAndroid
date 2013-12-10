package com.example.excelcreate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.content.Context;
import android.util.Log;
import android.view.Menu;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;
import android.widget.Toast;

public class MainActivity extends Activity {
	
	static String TAG = "ExelLog";
	Button btn1;
	 private static String FILE = Environment.getExternalStorageDirectory()+"/firstexcell.xls";
		
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		btn1=(Button)findViewById(R.id.button1);
		btn1.setOnClickListener(new OnClickListener() 
		{
			
			@Override
			public void onClick(View arg0) 
			{
				Boolean b=saveExcelFile(getApplicationContext(),FILE);
				if(b)
				{
					Toast.makeText(getApplicationContext(), "File Created", Toast.LENGTH_SHORT).show();
				}
				else
				{
					Toast.makeText(getApplicationContext(), "File Not Created Created", Toast.LENGTH_SHORT).show();
				}
				
			}
		});
	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(R.menu.main, menu);
		return true;
	}
	
	
	
	
	 private static boolean saveExcelFile(Context context, String fileName) { 
		 
	        // check if available and not read only 
	        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) { 
	            Log.e(TAG, "Storage not available or read only"); 
	            return false; 
	        } 
	 
	        boolean success = false; 
	 
	        //New Workbook
	        Workbook wb = new HSSFWorkbook();
	 
	        Cell c = null;
	 
	        //Cell style for header row
	        CellStyle cs = wb.createCellStyle();
	        cs.setFillForegroundColor(HSSFColor.LIME.index);
	        cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	        
	        //New Sheet
	        Sheet sheet1 = null;
	        sheet1 = wb.createSheet("myOrder");
	 
	        // Generate column headings
	        Row row = sheet1.createRow(0);
	 
	        c = row.createCell(0);
	        c.setCellValue("Item Number");
	        c.setCellStyle(cs);
	 
	        c = row.createCell(1);
	        c.setCellValue("Quantity");
	        c.setCellStyle(cs);
	 
	        c = row.createCell(2);
	        c.setCellValue("Price");
	        c.setCellStyle(cs);
	 
	        sheet1.setColumnWidth(0, (15 * 500));
	        sheet1.setColumnWidth(1, (15 * 500));
	        sheet1.setColumnWidth(2, (15 * 500));
	 
	        // Create a path where we will place our List of objects on external storage 
	        
	        File file=new File(FILE);
			if(!(file.exists()))
			{
				try {
					file.createNewFile();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
					Log.e("FileCreated", "filsesss");
				}
			}
	        
	        
	        
	        
	       // File file = new File(context.getExternalFilesDir(null), fileName); 
	        FileOutputStream os = null; 
	 
	        try { 
	            os = new FileOutputStream(file);
	            wb.write(os);
	            Log.w("FileUtils", "Writing file" + file); 
	            success = true; 
	        } catch (Exception e) { 
	            Log.w("FileUtils", "Error writing " + file, e); 
	        } finally { 
	            try { 
	                if (null != os) 
	                    os.close(); 
	            } catch (Exception ex) { 
	            } 
	        } 
	        return success; 
	    } 
	 
	 
	 
	 
	 
	 private static void readExcelFile(Context context, String filename) { 
		 
	        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) 
	        { 
	            Log.e(TAG, "Storage not available or read only"); 
	            return; 
	        } 
	 
	        try{
	            // Creating Input Stream 
	            File file = new File(context.getExternalFilesDir(null), filename); 
	            FileInputStream myInput = new FileInputStream(file);
	 
	            // Create a POIFSFileSystem object 
	            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
	 
	            // Create a workbook using the File System 
	            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);
	 
	            // Get the first sheet from workbook 
	            HSSFSheet mySheet = myWorkBook.getSheetAt(0);
	 
	            /** We now need something to iterate through the cells.**/
	            Iterator rowIter = mySheet.rowIterator();
	 
	            while(rowIter.hasNext()){
	                HSSFRow myRow = (HSSFRow) rowIter.next();
	                Iterator cellIter = myRow.cellIterator();
	                while(cellIter.hasNext()){
	                    HSSFCell myCell = (HSSFCell) cellIter.next();
	                    Log.d("Tag", "Cell Value: " +  myCell.toString());
	                    Toast.makeText(context, "cell Value: " + myCell.toString(), Toast.LENGTH_SHORT).show();
	                }
	            }
	        }catch (Exception e){e.printStackTrace(); }
	 
	        return;
	    } 
	
	
	 public static boolean isExternalStorageReadOnly() { 
	        String extStorageState = Environment.getExternalStorageState(); 
	        if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState)) { 
	            return true; 
	        } 
	        return false; 
	    } 
	 
	    public static boolean isExternalStorageAvailable() { 
	        String extStorageState = Environment.getExternalStorageState(); 
	        if (Environment.MEDIA_MOUNTED.equals(extStorageState)) { 
	            return true; 
	        } 
	        return false; 
	    } 
	 
	 

}
