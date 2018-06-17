package com.excelgo.go_hashmap;

import java.io.FileNotFoundException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class GoMain {

	public static void main(String[] args) throws GeneralSecurityException {	
		 try {  
				// �ļ�·��
	            String filepath = "D:\\MicrosoftExcel.xlsx";  
	            GoExcel excelReader = new GoExcel(filepath);  

				// �Զ�ȡExcel��������� ��ʹ��Map
	           Map<Integer,Object> mapTitle = excelReader.readExcelTitle2();  

	            // �Զ�ȡExcel������ݲ���  
               Map<Integer, Map<Integer,Object>> map = excelReader.readExcelContent();  

	    		//�������ϱ���Ա�����
	    		List<String> userids = new ArrayList<String>();
	               for(int i=1;i<=map.keySet().size();i++){
	       			//��������Ƿ���ڱ�ʶ����
	       			boolean isin = false;
	       			//�������ݶ���
	       			String kv = map.get(i).values().toArray()[0].toString();
	       			//�������б��
	    			for(int j=0;j<userids.size();j++){
	    				//�ж����ݶ������Ƿ��Ѿ����ڱ�ż���
	    				if(userids.get(j).equals(kv)){
	    					//��ʶ���ݶ������Ѿ����ڱ�ż���
	    					isin = true;
	    				}
	    			}
	    			//�жϱ�ʶ������ʶ���Ϊ�����ڣ������ݶ�������ӵ���ż���
	    			if(!isin){
	    				//�����ݶ�������ӵ���ż���
	    				userids.add(kv);
	    			}
	               }
	               
	               for(int i=0;i<userids.size();i++){
	            	   //�������󱣴浱ǰ�������ı��
	            	   String userid = userids.get(i);
	            	   
		            	 //�����ʼ������ֶ�
		       			String str1 = "";
		       			//�����ʼ������ֶ�
		       			String name = "";
		       			//�����ʼ��ռ���ַ
		       			String mail = "";
		       			
		       			for(int j=1;j<=map.keySet().size();j++){
		       				Object[] keyv = map.get(j).values().toArray();
		       				String kv = keyv[0].toString();
		       				
		       				//�ж����ݶ������Ƿ�͵�ǰ�������ı��һ�£������һ�������ݶ����ֵд���ʼ�����
		       				if(userid.equals(kv)){
		       					str1 += "<tr>";
		       					str1 += "<td width='100'>"+keyv[0].toString()+"</td>";//���
		       					str1 += "<td width='60'>"+keyv[1].toString()+"</td>";//����
		       					for(int k=3;k<keyv.length;k++){
		       						str1 += "<td width='80'>"+keyv[k].toString()+"</td>";
		       					}
		       					str1 += "</tr>";
		       					
		       					name = keyv[1].toString();
		       					mail = keyv[2].toString();
		       				}
		       			}
		       			
		       		//������Ϣ��ı�����
	      	         String str0 = name + "��<br>" 
	      	        		 	+"���ã����Ķ��˵����£�����ա�<br>"
	      	        		 	+ "<table border='1'>" 
	      	        		 	+ "<tr>"
	      	        		 	+ "<th>���</th>"
	      	        		 	+ "<th>����</th>";
	      	         for(int j=3;j<mapTitle.keySet().size();j++){
	      	        	 str0 += "<th>"+mapTitle.get(j)+"</th>";
	      	         }
	      	         str0 += "</tr>";
	      	         
	      	         String str = str0 + str1;
		       			
	      	         GoMail gomail = new GoMail();
	      	         gomail.goMail(name, mail, str);
	               }      
	        } catch (FileNotFoundException e) {  
	            System.out.println("δ�ҵ�ָ��·�����ļ�!");  
	            e.printStackTrace();  
	        }catch (Exception e) {  
	            e.printStackTrace();  
	        }  
	    }
}
