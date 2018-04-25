package com.hnthinker.mail;

import java.io.IOException;  
import java.util.List;  
  
import com.hnthinker.mail.XlsMain;  
  
public class Main {  
  
    public static void main(String[] args) {  
        //这个类主要是设置邮件     
        MailSenderInfo mailInfo = new MailSenderInfo();  
        mailInfo.setMailServerHost("smtp.exmail.qq.com");  
        mailInfo.setMailServerPort("25");  
        mailInfo.setValidate(true);  
        //设置发送邮箱和接收邮箱  
        mailInfo.setUserName("zhangmk@hnthinker.com");  
        mailInfo.setPassword("zhangmksw123");//您的邮箱密码      
        mailInfo.setFromAddress("zhangmk@hnthinker.com");  
//        mailInfo.setToAddress("***@qq.com");  
        //mailInfo.setToAddress("1291665093@qq.com");小岩邮箱  
        //mailInfo.setToAddress("772093950@qq.com"); 
        mailInfo.setSubject("测试标题");  
        mailInfo.setContent("设置邮箱内容 测试");  
        //这个类主要来发送邮件     
        SimpleMailSender sms = new SimpleMailSender();  
        /* 
         *  
        sms.sendTextMail(mailInfo);//发送文体格式     
        sms.sendHtmlMail(mailInfo);//发送html格式    
         */  
        //下面为读取excel数据然后发送  
        XlsMain xlsMain = new XlsMain();  
        try {  
            List<String []> list = xlsMain.readXls();  
            String [] title0 = list.get(0);//年月日  
            String [] title1 = list.get(1);  
            for(int i = 2; i < 3; i++){  
                String [] s = list.get(i); 
                mailInfo.setToAddress(s[2]);
                StringBuilder sb = new StringBuilder(); 
                sb.append(" <table background-color =\"gray\" width=\"400px\" border=\"5\">");
                sb.append("<tr><th align=\"center\" colspan=\"2\">" + title0[0] + "</th><tr/>"); //标题
                for(int j = 0; j < 19; j++){  
                    if(j == 2){  
                    	continue;
                    }
                    if(!title1[j].equals("0.0")){  //事项
                        sb.append("<tr><td>"+title1[j] + "</td>");  
                    }  
                    sb.append("<td>"+s[j] + "</td></tr>");  //值
                }
                sb.append("</table>");
                sb.append("<br/>本邮件由系统自动发送，实际工资发放以工资条为准！<br/>"); 
                sb.append("<br/>有疑问请联系：张明凯 zhangmk@hnthinker.com  0371-8888-888！<br/>");  
                sb.append("<br/><font color='red'>系统邮件，请勿直接回复！</font><br/>");  
                mailInfo.setSubject(s[1]  + "-工资信息");  
                mailInfo.setContent(sb.toString());  
                sms.sendHtmlMail(mailInfo);//发送文体格式     
            }  
        } catch (IOException e) {  
           e.printStackTrace();  
        }  
    }  
}  
