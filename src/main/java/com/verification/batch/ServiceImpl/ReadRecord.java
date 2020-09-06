/**
 * 
 */
package com.verification.batch.ServiceImpl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mysql.cj.jdbc.result.ResultSetMetaData;
import com.verification.batch.Batch;



/**
 *
 */
public class ReadRecord {

	/**
	 * @param args
	 * @throws Exception 
	 */
	public static void main(String[] args) throws Exception {
		
		ReadRecord r= new ReadRecord();
		r.readRecords("Recieve Payment");
		
	}
         public Batch readRecords(String batchn){
		 String jdbcURL = "jdbc:postgresql://localhost:5432/Report";
	        String username = "postgres";
	        String password = "avadakedavra";
	        Batch batch=new Batch();
	        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password)) {
	            
	            System.out.println("Connection established......");
	        	String sql = "select p.*,e.EVENTCOUNT,c.ERRORCOUNT from lis.tbbatch p  join lis.tbeventreport e on p.BATCH_ID=e.BATCH_ID  join lis.tberrorreport c on p.BATCH_ID=c.BATCH_ID  where p.BATCH_NAME='"+batchn+"'";
	          //  String sql = "select *  from lis.tbbatch"; 
	            Statement statement = connection.createStatement();
	            int count=0;
	            ResultSet result = statement.executeQuery(sql);
	            
	            while(result.next()){
	            	
	            	int batchid = result.getInt ("BATCH_ID");
	                String batchname = result.getString ("BATCH_NAME");
	                String status = result.getString ("STATUS");
	                int eventcount=result.getInt("EVENTCOUNT");
	                int errorcount=result.getInt("ERRORCOUNT");
	                
	                System.out.println (
	                        "BATCHID = " + batchid
	                        + ", BATCHNAME = " + batchname
	                        + ", STATUS = " + status
	                        + ", EVENTCOUNT = " + eventcount
	                        + ", ERRORCOUNT = " + errorcount);
	                ++count;
	          batch.setBatchId(batchid);
	          batch.setBatchName(batchname);
	          batch.setBatchStatus(status);
         batch.setEventcount(eventcount);
	          batch.setErrorcount(errorcount);
}
	            result.close();
	            statement.close();
	            
	            System.out.println (count + " rows were retrieved");
	           

	 
	        }
	        
	        catch (SQLException e) {
	            System.out.println("Datababse error:");
	            e.printStackTrace();
	        	}    
	        return batch;
	        } 
		
	}
	

	        
