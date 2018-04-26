package com.hnthinker.mail;


import java.awt.Button;
import java.awt.FileDialog;
import java.awt.Frame;
import java.awt.GridLayout;
import java.awt.Label;
import java.awt.Panel;
import java.awt.TextField;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.IOException;
import java.util.List;

	public class MainUi {
	    private Frame f;//窗体
	    String file; // 定义一个变量，存储文件路径以及名称
	    Label label1 = new Label("邮箱服务器");
	    TextField text1 = new TextField(20);
	    Label label2 = new Label("邮箱端口号");
	    TextField text2 = new TextField(20);
	    Label label3 = new Label("您的邮箱");
	    TextField text3 = new TextField(20);
	    Label label4 = new Label("邮箱密码");
	    TextField text4 = new TextField(20);
	    Label fileName =new Label("文件名称");
	    Button b1 = new Button("选择文件");
	    TextField fileChoose = new TextField(20);
	    FileDialog d1 = new FileDialog(f,"load file",FileDialog.LOAD);
	    Button b2 = new Button("确定");
	    Button b3 = new Button("取消");
	    // 构造方法
	    MainUi() {
	        // 控件初始化
	        initUI();
	    }
	    private void initUI() {
	        f = new Frame("工资单发送");
	        f.setSize(500, 400); // 设置窗口的宽高
	        f.setLocation(100, 100); // 设置窗口的起始点
	        f.setResizable(false); // 设置窗口一旦创建好，不能在改变大小。
	        GridLayout layout = new GridLayout(8,0);
	        f.setLayout(layout);
	        Panel p1 = new Panel();
	        p1.setSize(380, 50);
	        Panel p2 = new Panel();
	        p2.setSize(380, 50);
	        Panel p3 = new Panel();
	        p3.setSize(380, 50);
	        Panel p4 = new Panel();
	        p4.setSize(380, 50);
	        Panel p5 = new Panel();
	        p5.setSize(380, 50);
	        Panel p6 = new Panel();
	        p6.setSize(380, 50);
	        
	        
	        p1.add(label1);
	        text1.setText("smtp.exmail.qq.com");
	        p1.add(text1);
	        p2.add(label2);
	        text2.setText("25");
	        p2.add(text2);
	        p3.add(label3);
	        p3.add(text3);
	        p4.add(label4);
	        p4.add(text4);
	        p5.add(fileName);
	        p5.add(fileChoose);
	        p5.add(b1);
	        p6.add(b2);
	        p6.add(b3);
	        f.add(p1);
	        f.add(p2);
	        f.add(p3);
	        f.add(p4);
	        f.add(p5);
	        f.add(p6);
	        myEvent();
	        //f.pack();
	        f.setVisible(true);
	    }
	    private void myEvent(){ 
	    	//关闭窗口事件
	    	f.addWindowListener(new WindowAdapter() {
	            @Override
	            public void windowClosing(WindowEvent e)
	            {
	               System.exit(0);
	            }
	        });
	        //添加一个活动监听  
	        b1.addActionListener(new ActionListener() {  
	              
	            @Override  
	            public void actionPerformed(ActionEvent e) {  
	                System.out.println("选择文件");  
	                d1.setVisible(true);
	                file = d1.getDirectory()+d1.getFile();
	                fileChoose.setText(file);
	            }  
	        });  
	        b2.addActionListener(new ActionListener() {  
	              
	            @Override  
	            public void actionPerformed(ActionEvent e) {  
	                System.out.println("发送");  
	              //这个类主要是设置邮件     
	                MailSenderInfo mailInfo = new MailSenderInfo();  
	                mailInfo.setMailServerHost(text1.getText());  
	                mailInfo.setMailServerPort(text2.getText());  
	                mailInfo.setValidate(true);  
	                //设置发送邮箱和接收邮箱  
	                mailInfo.setUserName(text3.getText());  
	                mailInfo.setPassword(text4.getText());//您的邮箱密码      
	                mailInfo.setFromAddress(text3.getText());   
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
	                    List<String []> list = xlsMain.readXls(fileChoose.getText());  
	                    String [] title0 = list.get(0);//年月日  
	                    String [] title1 = list.get(1);  
	                    for(int i = 2; i < list.size()-2; i++){  
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
	                        sb.append("<br/><font color='red'>系统邮件，请勿直接回复！</font><br/>");  
	                        mailInfo.setSubject(s[1]  + "-工资信息");  
	                        mailInfo.setContent(sb.toString());  
	                        sms.sendHtmlMail(mailInfo);//发送文体格式     
	                    }  
	                } catch (IOException ex) {  
	                   ex.printStackTrace();  
	                }  
	            }  
	        });  
	        b3.addActionListener(new ActionListener() {  
	              
	            @Override  
	            public void actionPerformed(ActionEvent e) {  
	                System.out.println("取消发送");  
	                System.exit(0);
	            }  
	        });  
	          
	    }  
	    public static void main(String[] args) {
	        // TODO 自动生成的方法存根
	        new MainUi();
	    }

	}
