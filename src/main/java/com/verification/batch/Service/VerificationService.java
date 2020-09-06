package com.verification.batch.Service;

import java.io.FileNotFoundException;

import javax.mail.MessagingException;
import javax.mail.internet.AddressException;

public interface VerificationService {

	public void writeDataToExcel() throws FileNotFoundException, Exception  ;
	
	public void mailExcel() throws FileNotFoundException, AddressException, MessagingException;
}
