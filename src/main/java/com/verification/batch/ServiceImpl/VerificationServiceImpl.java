package com.verification.batch.ServiceImpl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.Month;
import java.util.Properties;
import java.util.logging.Level;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;

import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.util.HSSFColor;
//import org.apache.poi.hpsf.Date;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.verification.batch.Batch;
import com.verification.batch.Service.VerificationService;



public class VerificationServiceImpl implements VerificationService {

	
	@SuppressWarnings("resource")
	public void createHeaderCol(XSSFSheet sheet,XSSFWorkbook Resultworkbook, File result ) throws Exception
	{
//		File Result = new File("verification.xlsx");
//		FileInputStream ResultFile = new FileInputStream(Result);
//		XSSFWorkbook Resultworkbook = new XSSFWorkbook(ResultFile);
//		XSSFSheet resultSheet = Resultworkbook.getSheetAt(0);
		XSSFSheet resultSheet =sheet;
		
		//style 1 for header
		CellStyle style1 = Resultworkbook.createCellStyle();
		CreationHelper createHelper = Resultworkbook.getCreationHelper();  
		style1.setFillBackgroundColor(IndexedColors.ORANGE.getIndex());
		style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style1.setBorderBottom(BorderStyle.THIN);
		style1.setBorderLeft(BorderStyle.THIN);
		style1.setBorderTop(BorderStyle.THIN);
		style1.setBorderRight(BorderStyle.THIN);
		style1.setAlignment(HorizontalAlignment.CENTER);
		 Font font = Resultworkbook.createFont();
         font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
         style1.setFont(font);
     
		
		resultSheet.createRow(0).createCell(5).setCellValue("2LS PRODUCTION MONITORING VERIFICATION REPORT");
		resultSheet.getRow(0).setRowStyle(style1);
		//resultSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 12));
		resultSheet.createRow(1).createCell(0).setCellValue("DATE");
		
		CellStyle styleDate = Resultworkbook.createCellStyle();
		
		styleDate.setDataFormat(createHelper.createDataFormat().getFormat(
				"dd-mm-yyyy"));
		resultSheet.getRow(1).createCell(1).setCellStyle(styleDate);
		
		
		
//		  DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd");  
//		   LocalDateTime now = LocalDateTime.now();  
//		   System.out.println(dtf.format(now));
		
//		else if ( value instanceof Date ) {
//		    cell.setCellValue( (Date) value );
		
//		 Date date = Calendar.getInstance().getTime();  
//         DateFormat dateFormat = new SimpleDateFormat("yyyy-mm-dd hh:mm:ss");  
//         String strDate = dateFormat.format(date);  
//		    
		    
		//System.out.println(java.time.LocalDate.now());
			//resultSheet.getRow(1).createCell(1).setCellValue(java.time.LocalDate.now());
		resultSheet.getRow(1).getCell(1).setCellValue((Date) new Date());
		resultSheet.createRow(2).createCell(0).setCellValue("APPLICATION");
		//resultSheet.createRow(3).createCell(0).setCellValue("MESSEGING LOG VERIFICATON STATUS");
		//resultSheet.getRow(3).setRowStyle(style1);
		
		resultSheet.createRow(3).createCell(0).setCellValue("BATCH JOB VERIFICATION REPORT");
		
		//style 2 for header
		CellStyle style2 = Resultworkbook.createCellStyle();
		style2.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style2.setBorderBottom(BorderStyle.THIN);
		style2.setBorderLeft(BorderStyle.THIN);
		style2.setBorderTop(BorderStyle.THIN);
		style2.setBorderRight(BorderStyle.THIN);
		style2.setAlignment(HorizontalAlignment.CENTER);
		
		
		resultSheet.createRow(4).createCell(0).setCellValue("SR.NO");
		resultSheet.getRow(4).createCell(1).setCellValue("JOB NAME");
		resultSheet.getRow(4).createCell(2).setCellValue("BATCH");
		resultSheet.getRow(4).createCell(3).setCellValue("STATUS");
		resultSheet.getRow(4).createCell(4).setCellValue("event report exist?(number of event record)");
		resultSheet.getRow(4).createCell(5).setCellValue("error report exist?(number of error record)");
		resultSheet.getRow(4).createCell(6).setCellValue("PATH");
		resultSheet.getRow(4).createCell(7).setCellValue("Result");

		resultSheet.getRow(4).getCell(0).setCellStyle(style2);
		resultSheet.getRow(4).getCell(1).setCellStyle(style2);
		resultSheet.getRow(4).getCell(2).setCellStyle(style2);
		resultSheet.getRow(4).getCell(3).setCellStyle(style2);
		resultSheet.getRow(4).getCell(4).setCellStyle(style2);
		resultSheet.getRow(4).getCell(5).setCellStyle(style2);
		resultSheet.getRow(4).getCell(6).setCellStyle(style2);
		resultSheet.getRow(4).getCell(7).setCellStyle(style2);
		try {
			FileOutputStream out = new FileOutputStream(result);
			Resultworkbook.write(out);
			out.close();
			System.out.println("header written successfully..");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	
	
	
	public void writeDataToExcel() throws Exception {
		// TODO Auto-generated method stub
	VerificationServiceImpl impl=new VerificationServiceImpl();
	
    XSSFWorkbook workbook = new XSSFWorkbook(); 
    XSSFSheet sheet = workbook.createSheet(); 
    File result = new File("verification.xlsx");

	
	try {
		impl.createHeaderCol(sheet,workbook,result);
	} catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
		FileInputStream input = new FileInputStream(
				"C:\\Users\\kashe\\OneDrive\\Documents\\Work\\workspace\\Verification Report\\resources\\common.properties");
		System.out.println("file read batch");
		Properties config = new Properties();
		try {
			config.load(input);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		String[] dailyBatch = config.getProperty("dailyBatch").split(";");
		System.out.println(dailyBatch);
		VerificationServiceImpl write;
		int count=1;
		for(String batch:dailyBatch)
		{
			sheet.getLastRowNum();
			System.out.println(batch);
			write=new VerificationServiceImpl();
			write.writeExcel(batch,count,sheet,workbook,result);
			count++;
			
		}
		
		//String[] biweeklyBatch = config.getProperty("biweeklybatch").split(";");
		

		
		
	}

	//write data into excel
	public void writeExcel(String batch,int count,XSSFSheet sheet,XSSFWorkbook workbook, File result) throws Exception
	{
		FileInputStream input = new FileInputStream(
				"C:\\Users\\kashe\\OneDrive\\Documents\\Work\\workspace\\Verification Report\\resources\\common.properties");
		System.out.println("file read batch");
		Properties config = new Properties();	
        ReadRecord readRecord=new ReadRecord();
        config.load(input);
Batch bat= readRecord.readRecords(batch);
        // Create a blank sheet 
  
        CellStyle style1 = workbook.createCellStyle();
		style1.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
		
		CellStyle style2 = workbook.createCellStyle();
		style2.setFillBackgroundColor(IndexedColors.RED1.getIndex());
		
		CellStyle style3 = workbook.createCellStyle();
		style3.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.getIndex());
		
		CellStyle style4 = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setUnderline(Font.U_SINGLE);
		font.setColor(Font.COLOR_RED);
		style4.setFont(font);
		CreationHelper createHelper = workbook.getCreationHelper();
        
Row dataRow; 
		
//int rowCount = sheet.getLastRowNum();
System.out.println(count);
		dataRow = sheet.createRow(count+5);
		dataRow.createCell(0).setCellValue(count);
		dataRow.createCell(1).setCellValue(bat.getBatchId());
		dataRow.createCell(2).setCellValue(bat.getBatchName());
		dataRow.createCell(3).setCellValue(bat.getBatchStatus());
		if(bat.getBatchStatus().equals("OK")) dataRow.getCell(3).setCellStyle(style1);
		if(bat.getBatchStatus().equals("NOK")) dataRow.getCell(3).setCellStyle(style2);
		if(bat.getBatchStatus().equals("RUN")) dataRow.getCell(3).setCellStyle(style3);
		dataRow.createCell(4).setCellValue("yes ("+bat.getEventcount()+" records )");	
		dataRow.createCell(5).setCellValue("yes ("+bat.getErrorcount()+" records )");
		XSSFHyperlink link = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
		if (bat.getBatchName().equalsIgnoreCase("Omega Interface")) {
			dataRow.createCell(6).setCellValue(config.getProperty("OMEGA_INPUT_FILE"));
			
			//  link.setAddress("file:///C:/Users/kashe/OneDrive/Documents/Work/workspace/Verification Report/InputFile.txt");
		      
			
			 link.setAddress("https://www.google.com");
			 dataRow.getCell(6).setHyperlink((XSSFHyperlink) link);
		      dataRow.getCell(6).setCellStyle(style4);
		}
		else if (bat.getBatchName().equalsIgnoreCase("Claim Payment Bank Response")) {
			dataRow.createCell(6).setCellValue(config.getProperty("CLAIM_OUTPUT_FILE"));
		 link.setAddress("https://www.google.com");
			 dataRow.getCell(6).setHyperlink((XSSFHyperlink) link);
		      dataRow.getCell(6).setCellStyle(style4);
		}
		sheet.autoSizeColumn(1);
		sheet.autoSizeColumn(2);

        try { 
            // this Writes the workbook gfgcontribute 
            FileOutputStream out = new FileOutputStream(result); 
            workbook.write(out); 
            out.close(); 
            System.out.println(count+" written successfully on disk."); 
        } 
        catch (Exception e) { 
            e.printStackTrace(); 
        } 
	}
	
	
	public void mailExcel() throws FileNotFoundException, AddressException, MessagingException {
		// TODO Auto-generated method stub
		
		FileInputStream input = new FileInputStream(
				"D:\\java\\DemoHib\\resources\\config.properties");
		Properties config = new Properties();
		try {
			config.load(input);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		String host = config.getProperty("host");// host name
		String from = config.getProperty("username");// sender id
		String to = config.getProperty("reciever");// reciever id
		String pass = config.getProperty("password");// sender's password
		String fileAttachment = config.getProperty("fileAttachment");// file
		// name
		// for
		// attachment
		String port = config.getProperty("port");

		// system properties

		Properties prop = System.getProperties();

		// Setup mail server properties

		prop.put("mail.smtp.host", host);
		prop.put("mail.smtp.starttls.enable", "true");
		prop.put("mail.smtp.user", from);
		prop.put("mail.smtp.password", pass);
		prop.put("mail.smtp.port", port);
		prop.put("mail.smtp.auth", "true");

		// session

		Session session = Session.getInstance(prop, null);

		// Define message
		try {

			Month month = this.getCurrentMonth();

			MimeMessage message = new MimeMessage(session);
			message.setFrom(new InternetAddress(from));
			message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));
			message.setSubject("Timesheet attachment for month " + month);
		String htmlLink= "<a href=D:\\> Google </a> </br>";
			//message.setContent(htmlLink, "text/html");

			// create the message part

			MimeBodyPart messageBodyPart = new MimeBodyPart();
			//String htmlContent = "<a href=www.google.com>google</a>";

			// message body

			String[] name = from.split("\\.");

			String[] surname = name[1].split("@");
			//messageBodyPart.setContent(htmlContent, "text/html");
			//messageBodyPart.setText(htmlLink+"\n\nThanks & Regards, \n" + name[0] + " " + surname[0]);
			messageBodyPart.setText(htmlLink);

			Multipart multipart = new MimeMultipart();
			multipart.addBodyPart(messageBodyPart);

			// attachment

			//messageBodyPart = new MimeBodyPart();
			DataSource source = new FileDataSource(config.getProperty("filePath"));
			messageBodyPart.setDataHandler(new DataHandler(source));
			messageBodyPart.setFileName(fileAttachment);
			multipart.addBodyPart(messageBodyPart);
			message.setContent(multipart,"text/html");

			// send message to reciever

			Transport transport = session.getTransport("smtp");
			transport.connect(host, from, pass);
			transport.sendMessage(message, message.getAllRecipients());
			System.out.println("Mail Sent Successfully!");
			transport.close();

		}

		catch (javax.mail.AuthenticationFailedException e) {
			java.util.logging.Logger.getLogger("Test").log(Level.INFO, "BOOM!", e);
		}
	}
		public Month getCurrentMonth() {

			
			LocalDate currentdate = LocalDate.now().plusMonths(-1);
			Month currentMonth = currentdate.getMonth();
			System.out.println("Current month: " + currentMonth);

			return currentMonth;
		}
		
		public void addHyperlink() {
		
		}
	}



