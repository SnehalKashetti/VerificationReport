package com.verification.batch;

import java.io.FileNotFoundException;

import javax.mail.MessagingException;
import javax.mail.internet.AddressException;

import com.verification.batch.ServiceImpl.VerificationServiceImpl;

public class MainVerification {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
VerificationServiceImpl vsi = new VerificationServiceImpl();

try {
	vsi.writeDataToExcel();
} catch (FileNotFoundException e1) {
	// TODO Auto-generated catch block
	e1.printStackTrace();
}
//try {
//	vsi.mailExcel();
//} catch (AddressException e) {
//	// TODO Auto-generated catch block
//	e.printStackTrace();
//} catch (FileNotFoundException e) {
//	// TODO Auto-generated catch block
//	e.printStackTrace();
//} catch (MessagingException e) {
//	// TODO Auto-generated catch block
//	e.printStackTrace();
//}
	}

}
