package com.excelgo.go_hashmap;  

import java.security.GeneralSecurityException;
import java.util.Properties;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import com.sun.mail.util.MailSSLSocketFactory;
 
public class GoMail{
	
	public void goMail(String name, String mail, String str) throws GeneralSecurityException{
	  // �ռ��˵�������
		String to = mail;
	 
	      // �����˵�������
		String from = "example@qq.com";
	 
	      // ָ�������ʼ�������Ϊsmtp.qq.com
		String host = "smtp.exmail.qq.com";
	 
	      // ��ȡϵͳ����
		Properties properties = System.getProperties();
	 
	      // �����ʼ�������
		properties.setProperty("mail.smtp.host", host);     
		properties.put("mail.smtp.auth", "true");
	  
		  //QQ�����SSL����
		MailSSLSocketFactory sf = new MailSSLSocketFactory();
		sf.setTrustAllHosts(true);
	    properties.put("mail.smtp.ssl.enable", "true");
	 	properties.put("mail.smtp.ssl.socketFactory", sf);
			 
	      // ��ȡĬ��session����
	 	Session session = Session.getDefaultInstance(properties,new Authenticator(){
	 		public PasswordAuthentication getPasswordAuthentication()
	 		{
	 			return new PasswordAuthentication("example@qq.com, "password"); //�������ʼ��û���������
	 		}
	 	});
	 	try{
	         // ����Ĭ�ϵ� MimeMessage ����
	         MimeMessage message = new MimeMessage(session);
	 
	         // Set From: ͷ��ͷ�ֶ�
	         message.setFrom(new InternetAddress(from));
	 
	         // Set To: ͷ��ͷ�ֶ�
	         message.addRecipient(Message.RecipientType.TO,
	                                  new InternetAddress(to));
	 
	         // Set Subject: ͷ��ͷ�ֶ�
	         message.setSubject("���˵�֪ͨ");

	         
	         str+="</table>";
	         
	         message.setContent( str ,"text/html;charset=gb2312" );
	 
	         // ������Ϣ
	         Transport.send(message);
	         System.out.println("�ʼ����ͳɹ���");
	      }catch (MessagingException mex) {
	         mex.printStackTrace();
	      }
	}

      
}
