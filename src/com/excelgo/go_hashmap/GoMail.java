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
	  // 收件人电子邮箱
		String to = mail;
	 
	      // 发件人电子邮箱
		String from = "example@qq.com";
	 
	      // 指定发送邮件的主机为smtp.qq.com
		String host = "smtp.exmail.qq.com";
	 
	      // 获取系统属性
		Properties properties = System.getProperties();
	 
	      // 设置邮件服务器
		properties.setProperty("mail.smtp.host", host);     
		properties.put("mail.smtp.auth", "true");
	  
		  //QQ邮箱的SSL加密
		MailSSLSocketFactory sf = new MailSSLSocketFactory();
		sf.setTrustAllHosts(true);
	    properties.put("mail.smtp.ssl.enable", "true");
	 	properties.put("mail.smtp.ssl.socketFactory", sf);
			 
	      // 获取默认session对象
	 	Session session = Session.getDefaultInstance(properties,new Authenticator(){
	 		public PasswordAuthentication getPasswordAuthentication()
	 		{
	 			return new PasswordAuthentication("example@qq.com, "password"); //发件人邮件用户名、密码
	 		}
	 	});
	 	try{
	         // 创建默认的 MimeMessage 对象
	         MimeMessage message = new MimeMessage(session);
	 
	         // Set From: 头部头字段
	         message.setFrom(new InternetAddress(from));
	 
	         // Set To: 头部头字段
	         message.addRecipient(Message.RecipientType.TO,
	                                  new InternetAddress(to));
	 
	         // Set Subject: 头部头字段
	         message.setSubject("对账单通知");

	         
	         str+="</table>";
	         
	         message.setContent( str ,"text/html;charset=gb2312" );
	 
	         // 发送消息
	         Transport.send(message);
	         System.out.println("邮件发送成功！");
	      }catch (MessagingException mex) {
	         mex.printStackTrace();
	      }
	}

      
}
