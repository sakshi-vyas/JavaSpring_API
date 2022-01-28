package com.example.Controller;

import java.io.File;
import javax.mail.*;  
import javax.mail.internet.*;
import javax.mail.Session;
import javax.activation.*;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import java.util.Timer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.scheduling.annotation.Async;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import com.example.FirstProjectApplication;
import com.example.Pojo.Student;
import com.sun.mail.smtp.SMTPTransport;
@RestController 
public class MyController {
	
	// List<String> to_mails= new ArrayList<>();
		
	@GetMapping("/StudentInfo")
	public List<Student> getInfo() throws Exception {
		 String FILE_NAME = "./src/main/java/com/example/Controller/Project01.xlsx";
	        List<Student> list= new ArrayList<>();
		 try {

	            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
	            Workbook workbook = new XSSFWorkbook(excelFile);
	            Sheet datatypeSheet = workbook.getSheetAt(0);
	            Iterator<Row> iterator = datatypeSheet.iterator();

	    
	            Student student ;
	            
	            while (iterator.hasNext()) {
	            	 student = new Student();
	                Row currentRow = iterator.next();
	                Iterator<Cell> cellIterator = currentRow.iterator();
	                int i=0;
	                while (cellIterator.hasNext()) {

	                    Cell currentCell = cellIterator.next();
	                    if (i==0){
	                        student.setName(currentCell.getStringCellValue());
	                        } 
	                    if (i==1){
	                        student.setMail(currentCell.getStringCellValue());
	                        sendMail(student.getMail(),student.getName());
	                        //to_mails.add(student.getMail());
	                        } 
	                    if (i==2){
	                        student.setContact((long)currentCell.getNumericCellValue());
	                        } 
	                    i+=1;
	                }
	                list.add(student);
	            }
	        } catch (FileNotFoundException e) {
	            e.printStackTrace();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
		 System.out.println("read excel "+ new Date());
		 return list;
	}

	//@Scheduled(initialDelay=3000,fixedDelay=5000)
	//@Async
	public void sendMail(String mail,String name) throws Exception {
	System.out.println("mailing starts.. "+ Thread.currentThread().getName());
	        Properties prop = System.getProperties();
	        prop.put("mail.smtp.auth", "true");
	        prop.put("mail.smtp.starttls.enable", "true");
	        prop.put("mail.smtp.host", "smtp.gmail.com"); 
	        prop.put("mail.smtp.port", "587");
	        String username = "dummy216.mail@gmail.com";
	        String password = "dummy@123";
	     
	        Session session = Session.getDefaultInstance(prop, 
	                            new Authenticator(){
	                               protected PasswordAuthentication getPasswordAuthentication() {
	                                  return new PasswordAuthentication(username, password);
	                               }}); 
	        Message message;
//	        for (int i = 0; i < to_mails.size(); i++) {
	        message = prepareMessage(session,username,mail,name); 
	        Transport.send(message);
	        System.out.println("message sent successfully");
//	        }
	        
	       
	       
	}  
	        private static Message prepareMessage(Session session, String username,String mail,String name) {
	      
	      try{  
	    	  	Message message =new MimeMessage(session);
	         message.setFrom(new InternetAddress(username));  
	         message.setRecipient(Message.RecipientType.TO,new InternetAddress(mail));  
	         message.setSubject("Congratulations!!");  
	         message.setText("Dear "+name+"\r\n"
	         		+ "\r\n"
	         		+ "I hope you are doing well.\r\n"
	         		+ "\r\n"
	         		+ "I am writing this letter to make sure my brightest student gets a good amount of appreciation. "
	         		+ "I see that you have been a good student throughout the year and you must receive feedback regarding the same.\r\n"
	         		+ "\r\n"
	         		+ "I am impressed by all your work which you have submitted till now, and I must acknowledge the amount "
	         		+ "of dedication you have put in to finish them. I am amused by the time management methods you use and I must say,"
	         		+ " you can help out other students too with the same."
	         		+ " I find you really helpful and patient when you solve other student’s problems. It is really kind of you.");
	         return message;
	      }
	      catch (Exception ex)
	      {
	    	  ex.printStackTrace();
	      } 
	      return null;
	        }
	        
	      
	
}

