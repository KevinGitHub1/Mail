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
	    private Frame f;//����
	    String file; // ����һ���������洢�ļ�·���Լ�����
	    Label label1 = new Label("���������");
	    TextField text1 = new TextField(20);
	    Label label2 = new Label("����˿ں�");
	    TextField text2 = new TextField(20);
	    Label label3 = new Label("��������");
	    TextField text3 = new TextField(20);
	    Label label4 = new Label("��������");
	    TextField text4 = new TextField(20);
	    Label fileName =new Label("�ļ�����");
	    Button b1 = new Button("ѡ���ļ�");
	    TextField fileChoose = new TextField(20);
	    FileDialog d1 = new FileDialog(f,"load file",FileDialog.LOAD);
	    Button b2 = new Button("ȷ��");
	    Button b3 = new Button("ȡ��");
	    // ���췽��
	    MainUi() {
	        // �ؼ���ʼ��
	        initUI();
	    }
	    private void initUI() {
	        f = new Frame("���ʵ�����");
	        f.setSize(500, 400); // ���ô��ڵĿ��
	        f.setLocation(100, 100); // ���ô��ڵ���ʼ��
	        f.setResizable(false); // ���ô���һ�������ã������ڸı��С��
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
	    	//�رմ����¼�
	    	f.addWindowListener(new WindowAdapter() {
	            @Override
	            public void windowClosing(WindowEvent e)
	            {
	               System.exit(0);
	            }
	        });
	        //���һ�������  
	        b1.addActionListener(new ActionListener() {  
	              
	            @Override  
	            public void actionPerformed(ActionEvent e) {  
	                System.out.println("ѡ���ļ�");  
	                d1.setVisible(true);
	                file = d1.getDirectory()+d1.getFile();
	                fileChoose.setText(file);
	            }  
	        });  
	        b2.addActionListener(new ActionListener() {  
	              
	            @Override  
	            public void actionPerformed(ActionEvent e) {  
	                System.out.println("����");  
	              //�������Ҫ�������ʼ�     
	                MailSenderInfo mailInfo = new MailSenderInfo();  
	                mailInfo.setMailServerHost(text1.getText());  
	                mailInfo.setMailServerPort(text2.getText());  
	                mailInfo.setValidate(true);  
	                //���÷�������ͽ�������  
	                mailInfo.setUserName(text3.getText());  
	                mailInfo.setPassword(text4.getText());//������������      
	                mailInfo.setFromAddress(text3.getText());   
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
	                    List<String []> list = xlsMain.readXls(fileChoose.getText());  
	                    String [] title0 = list.get(0);//������  
	                    String [] title1 = list.get(1);  
	                    for(int i = 2; i < list.size()-2; i++){  
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
	                        sb.append("<br/><font color='red'>ϵͳ�ʼ�������ֱ�ӻظ���</font><br/>");  
	                        mailInfo.setSubject(s[1]  + "-������Ϣ");  
	                        mailInfo.setContent(sb.toString());  
	                        sms.sendHtmlMail(mailInfo);//���������ʽ     
	                    }  
	                } catch (IOException ex) {  
	                   ex.printStackTrace();  
	                }  
	            }  
	        });  
	        b3.addActionListener(new ActionListener() {  
	              
	            @Override  
	            public void actionPerformed(ActionEvent e) {  
	                System.out.println("ȡ������");  
	                System.exit(0);
	            }  
	        });  
	          
	    }  
	    public static void main(String[] args) {
	        // TODO �Զ����ɵķ������
	        new MainUi();
	    }

	}
