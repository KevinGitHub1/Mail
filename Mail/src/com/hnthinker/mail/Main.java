package com.hnthinker.mail;

import java.io.IOException;  
import java.util.List;  
  
import com.hnthinker.mail.XlsMain;  
  
public class Main {  
  
    public static void main(String[] args) {  
        //�������Ҫ�������ʼ�     
        MailSenderInfo mailInfo = new MailSenderInfo();  
        mailInfo.setMailServerHost("smtp.exmail.qq.com");  
        mailInfo.setMailServerPort("25");  
        mailInfo.setValidate(true);  
        //���÷�������ͽ�������  
        mailInfo.setUserName("zhangmk@hnthinker.com");  
        mailInfo.setPassword("zhangmksw123");//������������      
        mailInfo.setFromAddress("zhangmk@hnthinker.com");  
//        mailInfo.setToAddress("***@qq.com");  
        //mailInfo.setToAddress("1291665093@qq.com");С������  
        //mailInfo.setToAddress("772093950@qq.com"); 
        mailInfo.setSubject("���Ա���");  
        mailInfo.setContent("������������ ����");  
        //�������Ҫ�������ʼ�     
        SimpleMailSender sms = new SimpleMailSender();  
        /* 
         *  
        sms.sendTextMail(mailInfo);//���������ʽ     
        sms.sendHtmlMail(mailInfo);//����html��ʽ    
         */  
        //����Ϊ��ȡexcel����Ȼ����  
        XlsMain xlsMain = new XlsMain();  
        try {  
            List<String []> list = xlsMain.readXls();  
            String [] title0 = list.get(0);//������  
            String [] title1 = list.get(1);  
            for(int i = 2; i < 3; i++){  
                String [] s = list.get(i); 
                mailInfo.setToAddress(s[2]);
                StringBuilder sb = new StringBuilder(); 
                sb.append(" <table background-color =\"gray\" width=\"400px\" border=\"5\">");
                sb.append("<tr><th align=\"center\" colspan=\"2\">" + title0[0] + "</th><tr/>"); //����
                for(int j = 0; j < 19; j++){  
                    if(j == 2){  
                    	continue;
                    }
                    if(!title1[j].equals("0.0")){  //����
                        sb.append("<tr><td>"+title1[j] + "</td>");  
                    }  
                    sb.append("<td>"+s[j] + "</td></tr>");  //ֵ
                }
                sb.append("</table>");
                sb.append("<br/>���ʼ���ϵͳ�Զ����ͣ�ʵ�ʹ��ʷ����Թ�����Ϊ׼��<br/>"); 
                sb.append("<br/>����������ϵ�������� zhangmk@hnthinker.com  0371-8888-888��<br/>");  
                sb.append("<br/><font color='red'>ϵͳ�ʼ�������ֱ�ӻظ���</font><br/>");  
                mailInfo.setSubject(s[1]  + "-������Ϣ");  
                mailInfo.setContent(sb.toString());  
                sms.sendHtmlMail(mailInfo);//���������ʽ     
            }  
        } catch (IOException e) {  
           e.printStackTrace();  
        }  
    }  
}  
