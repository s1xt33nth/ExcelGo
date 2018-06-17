package com.excelgo.go_hashmap;

import java.io.FileNotFoundException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class GoMain {

	public static void main(String[] args) throws GeneralSecurityException {	
		 try {  
				// 文件路径
	            String filepath = "D:\\MicrosoftExcel.xlsx";  
	            GoExcel excelReader = new GoExcel(filepath);  

				// 对读取Excel表格标题测试 ，使用Map
	           Map<Integer,Object> mapTitle = excelReader.readExcelTitle2();  

	            // 对读取Excel表格内容测试  
               Map<Integer, Map<Integer,Object>> map = excelReader.readExcelContent();  

	    		//创建集合保存员工编号
	    		List<String> userids = new ArrayList<String>();
	               for(int i=1;i<=map.keySet().size();i++){
	       			//创建编号是否存在标识对象
	       			boolean isin = false;
	       			//创建数据对象
	       			String kv = map.get(i).values().toArray()[0].toString();
	       			//遍历现有编号
	    			for(int j=0;j<userids.size();j++){
	    				//判断数据对象编号是否已经存在编号集合
	    				if(userids.get(j).equals(kv)){
	    					//标识数据对象编号已经存在编号集合
	    					isin = true;
	    				}
	    			}
	    			//判断标识，若标识编号为不存在，则将数据对象编号添加到编号集合
	    			if(!isin){
	    				//将数据对象编号添加到编号集合
	    				userids.add(kv);
	    			}
	               }
	               
	               for(int i=0;i<userids.size();i++){
	            	   //创建对象保存当前被遍历的编号
	            	   String userid = userids.get(i);
	            	   
		            	 //创建邮件内容字段
		       			String str1 = "";
		       			//创建邮件姓名字段
		       			String name = "";
		       			//创建邮件收件地址
		       			String mail = "";
		       			
		       			for(int j=1;j<=map.keySet().size();j++){
		       				Object[] keyv = map.get(j).values().toArray();
		       				String kv = keyv[0].toString();
		       				
		       				//判断数据对象编号是否和当前被遍历的编号一致，若编号一致则将数据对象的值写入邮件内容
		       				if(userid.equals(kv)){
		       					str1 += "<tr>";
		       					str1 += "<td width='100'>"+keyv[0].toString()+"</td>";//编号
		       					str1 += "<td width='60'>"+keyv[1].toString()+"</td>";//姓名
		       					for(int k=3;k<keyv.length;k++){
		       						str1 += "<td width='80'>"+keyv[k].toString()+"</td>";
		       					}
		       					str1 += "</tr>";
		       					
		       					name = keyv[1].toString();
		       					mail = keyv[2].toString();
		       				}
		       			}
		       			
		       		//设置消息体的表格标题
	      	         String str0 = name + "：<br>" 
	      	        		 	+"您好，您的对账单如下，请查收。<br>"
	      	        		 	+ "<table border='1'>" 
	      	        		 	+ "<tr>"
	      	        		 	+ "<th>编号</th>"
	      	        		 	+ "<th>姓名</th>";
	      	         for(int j=3;j<mapTitle.keySet().size();j++){
	      	        	 str0 += "<th>"+mapTitle.get(j)+"</th>";
	      	         }
	      	         str0 += "</tr>";
	      	         
	      	         String str = str0 + str1;
		       			
	      	         GoMail gomail = new GoMail();
	      	         gomail.goMail(name, mail, str);
	               }      
	        } catch (FileNotFoundException e) {  
	            System.out.println("未找到指定路径的文件!");  
	            e.printStackTrace();  
	        }catch (Exception e) {  
	            e.printStackTrace();  
	        }  
	    }
}
