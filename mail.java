import java.awt.Color;
import java.awt.Component;
import java.awt.Cursor;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

import com.sun.mail.pop3.POP3Store;

import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JSpinner;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.SwingWorker;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.joda.time.DateTime.Property;


import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.Flags.Flag;
import javax.mail.internet.*;


import com.sun.mail.imap.IMAPFolder;
import com.sun.mail.imap.IMAPMessage;
import java.io.*;
import java.util.*;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.mail.search.FlagTerm;
import javax.swing.JTextPane;
import java.awt.Panel; 

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import javax.swing.JCheckBox;
import java.awt.event.ItemListener;
import java.awt.event.ItemEvent;
import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;





public class mail {

	
	private JFrame frame;
	private static JTable table;
	private JScrollPane scrollPane;

	public static JSpinner spinner;
	
	public static JSpinner spinner_1;
	
	public static int hour_j;
	
	public static int min_j;
	
	public static InputStream inp;

	public static Workbook wb;

	public static Sheet sheet ;

	public static Row row;

	public static Cell cell;
	
	public static  DefaultTableModel model2;
	
	public static String username;
	
	
	public static String password;
	
	public static String ip;
	
	public static String port;
	
	public static String table_proxy;
	
	public static Properties prop ;
	public static int delay;
	private JLabel label_2;
	private JLabel label_1;
	private static JLabel label_3;
	private static JLabel lblScheduleTasks;
	private static JLabel lblClickExecute;
	private static JLabel label_4;
	public static boolean imported ;
	public static String path;
	 public static  Properties mailServerProperties;
		public static Session getMailSession;
		public static MimeMessage generateMailMessage;
		private JLabel lblSelectSpecificTask;
		private static JTextField textField_1;
		private static JTextField textField_2;
		private static JLabel lblTo;
		
		public static double ID ;
		private static Panel panel_1;
		private JLabel lblDashboard_1;
		private JLabel lblSuccesssfullLogins;
		private JLabel lblFailedLogins;
		private JLabel lblTotalSpamMessage;
		private JLabel lblTotalUnreadMessaged;
		protected static boolean Reply;
		private static  JTextField textField;
		private static JTextField textField_3;
		private static JTextField textField_4;
		private static  JTextField textField_5;
		public static int s_l;
		public static int f_l;
		public static int t_s;
		public static int t_r;
		public static int t_u_r;
		
		public static JLabel label ;
		
		static SheetsQuickstart ev = new SheetsQuickstart();
		private static boolean userN;
		private static JTextField textField_6;
		public static JComboBox comboBox;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					
					userN = false;
					String Url_st = "https://sheets.googleapis.com/v4/spreadsheets/1svMRxiC1X6J9gWHxp02seDH0JXCc6mfh46QieLKAG24/values/Sheet1!A1:D5?key=AIzaSyC4Bxi7iVxYjfDhUT5K5c0hEth9CAbCvSc";

					if (Url_st.contains(" "))
						Url_st = Url_st.replaceAll(" ", "%20");

					System.out.println(Url_st);

					URL url = new URL(Url_st);

					HttpURLConnection conn = (HttpURLConnection) url.openConnection();

					conn.setRequestMethod("GET");

					int responsecode = conn.getResponseCode();

					if (responsecode != 200)
						throw new RuntimeException("HttpResponseCode: " + responsecode);
					else {

					}

					Scanner sc = new Scanner(url.openStream());
					String inline = "";
					while (sc.hasNext()) {
						inline += sc.nextLine();
					}
					System.out.println("\nJSON data in string format");
					//System.out.println(inline);
					sc.close();

					JSONParser parse = new JSONParser();

					JSONObject jobj = (JSONObject) parse.parse(inline);

					JSONArray jsonarr_1 = (JSONArray) jobj.get("values");

					// Get data for Results array
					for (int i = 0; i < jsonarr_1.size(); i++) {
						
						System.out.println("\n " +  jsonarr_1.get(i));
						
						String f = String.valueOf(jsonarr_1.get(i));
						
						f= f.replaceAll("\"", "");
						f= f.replaceAll("\\[", "");
						f= f.replaceAll("\\]", "");
						
						String [] arw = f.split(",");
						
						if(arw[0].equals("trader100") && arw[1].equals("TRUE"))
						{
								System.out.println("user has been authenticated!");
								userN=true;
						}
						// Store the JSON objects in an array
						// Get the index of the JSON object and print the values
						// as per the index
						/*
						JSONObject jsonobj_1 = (JSONObject) jsonarr_1.get(i);
						System.out.println("Elements under results array");
						System.out.println("\n " + jsonobj_1.toString());
						System.out.println("\nName: " + jsonobj_1.get("name"));
						System.out.println("Address: " + jsonobj_1.get("formatted_address"));

						*/
						
					
						}
					Reply= true;
					DayOfWeek dayOfWeek2 = DayOfWeek.from(LocalDate.now());
					
	        		
	        		String d = String.valueOf(dayOfWeek2);
	        	//	System.out.println(d);
	        		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH:mm");  
		        	   LocalDateTime now = LocalDateTime.now(); 
		        	   String da=dtf.format(now);
		        	 //  System.out.println(da);  
		        	   
		        	  // textField_5.setText(d+" "+da);
				
		        	   System.out.println(d+" "+da);
		        	   
		        	   String [] ho = da.split(":");
		        	   String h =ho[0];
		        	   String m =ho[1];
					if(userN == true)
					{
					mail window = new mail();
					window.frame.setVisible(true);
					
DateTime dt = new DateTime();
					
					Property month = dt.dayOfMonth();     // gets the current month
					int hours = dt.getHourOfDay(); // gets hour of day
					
					int min = dt.getMinuteOfHour();
					
					int seconds = dt.getSecondOfMinute();
System.out.println("hour is: "+hours+" minute is: " +min +" "+seconds);
					
					spinner.setValue(Integer.parseInt(h));
					
					spinner_1.setValue(Integer.parseInt(m));
					
					
					
				
					
				
					
				
					table_proxy="";
					
					 s_l=0;
					 f_l=0;
					 t_s=0;
				     t_r=0;
				     t_u_r=0;
				     bet2();
				     
				     
				     
				   
					}else{
						 JOptionPane.showMessageDialog(null, "NO internet connection or something is not right..");
					}
				          
				          
				      
				     
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public mail() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.getContentPane().setBackground(Color.BLACK);
		frame.setBounds(100, 100, 1045, 514);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		frame.setResizable(false);
		scrollPane = new JScrollPane();
		scrollPane.setBounds(182, 37, 500, 437);
		frame.getContentPane().add(scrollPane);
		
		JCheckBox chckbxDoNotReply = new JCheckBox("DO NOT REPLY");
		chckbxDoNotReply.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				JCheckBox chckbxNewCheckBox_1 = (JCheckBox) e.getSource();

                // The item affected by the event.
                Object item = e.getItem();

              

                if (e.getStateChange() == ItemEvent.SELECTED) {
                	//textField.setText(item.toString());
                	
                	System.out.println("checked");
                	
                	Reply = false;
                	

					
					 
                }

                if (e.getStateChange() == ItemEvent.DESELECTED) {
                  //  textArea.setText(item.toString() + " deselected.");
                	
                	System.out.println("not checked");
                	
                	Reply = true;
                }
			}
		});
		chckbxDoNotReply.setFont(new Font("Times New Roman", Font.BOLD, 12));
		chckbxDoNotReply.setHorizontalAlignment(SwingConstants.CENTER);
		chckbxDoNotReply.setBackground(Color.YELLOW);
		chckbxDoNotReply.setBounds(25, 299, 143, 23);
		frame.getContentPane().add(chckbxDoNotReply);
		
		panel_1 = new Panel();
		panel_1.setBackground(Color.WHITE);
		panel_1.setBounds(688, 37, 341, 437);
		frame.getContentPane().add(panel_1);
		panel_1.setLayout(null);
		
		lblDashboard_1 = new JLabel("DASHBOARD");
		lblDashboard_1.setFont(new Font("Times New Roman", Font.BOLD, 15));
		lblDashboard_1.setForeground(Color.YELLOW);
		lblDashboard_1.setBackground(Color.BLACK);
		lblDashboard_1.setOpaque(true);
		lblDashboard_1.setHorizontalAlignment(SwingConstants.CENTER);
		lblDashboard_1.setBounds(0, 26, 341, 21);
		panel_1.add(lblDashboard_1);
		
		lblSuccesssfullLogins = new JLabel("SUCCESSSFULL LOGINS:");
		lblSuccesssfullLogins.setFont(new Font("Times New Roman", Font.BOLD, 11));
		lblSuccesssfullLogins.setOpaque(true);
		lblSuccesssfullLogins.setHorizontalAlignment(SwingConstants.CENTER);
		lblSuccesssfullLogins.setForeground(Color.YELLOW);
		lblSuccesssfullLogins.setBackground(Color.BLACK);
		lblSuccesssfullLogins.setBounds(0, 75, 258, 21);
		panel_1.add(lblSuccesssfullLogins);
		
		lblFailedLogins = new JLabel("FAILED LOGINS:");
		lblFailedLogins.setFont(new Font("Times New Roman", Font.BOLD, 11));
		lblFailedLogins.setOpaque(true);
		lblFailedLogins.setHorizontalAlignment(SwingConstants.CENTER);
		lblFailedLogins.setForeground(Color.YELLOW);
		lblFailedLogins.setBackground(Color.BLACK);
		lblFailedLogins.setBounds(0, 127, 258, 21);
		panel_1.add(lblFailedLogins);
		
		lblTotalSpamMessage = new JLabel("TOTAL SPAM MESSAGES MOVED TO INBOX:");
		lblTotalSpamMessage.setFont(new Font("Times New Roman", Font.BOLD, 11));
		lblTotalSpamMessage.setOpaque(true);
		lblTotalSpamMessage.setHorizontalAlignment(SwingConstants.CENTER);
		lblTotalSpamMessage.setForeground(Color.YELLOW);
		lblTotalSpamMessage.setBackground(Color.BLACK);
		lblTotalSpamMessage.setBounds(0, 176, 258, 21);
		panel_1.add(lblTotalSpamMessage);
		
		lblTotalUnreadMessaged = new JLabel("TOTAL UNREAD MESSAGES FOUND:");
		lblTotalUnreadMessaged.setFont(new Font("Times New Roman", Font.BOLD, 11));
		lblTotalUnreadMessaged.setOpaque(true);
		lblTotalUnreadMessaged.setHorizontalAlignment(SwingConstants.CENTER);
		lblTotalUnreadMessaged.setForeground(Color.YELLOW);
		lblTotalUnreadMessaged.setBackground(Color.BLACK);
		lblTotalUnreadMessaged.setBounds(0, 235, 258, 21);
		panel_1.add(lblTotalUnreadMessaged);
		
		textField = new JTextField();
		textField.setHorizontalAlignment(SwingConstants.CENTER);
		textField.setEditable(false);
		textField.setBounds(284, 75, 47, 20);
		panel_1.add(textField);
		textField.setColumns(10);
		
		textField_3 = new JTextField();
		textField_3.setHorizontalAlignment(SwingConstants.CENTER);
		textField_3.setEditable(false);
		textField_3.setColumns(10);
		textField_3.setBounds(284, 127, 47, 20);
		panel_1.add(textField_3);
		
		textField_4 = new JTextField();
		textField_4.setHorizontalAlignment(SwingConstants.CENTER);
		textField_4.setEditable(false);
		textField_4.setColumns(10);
		textField_4.setBounds(284, 176, 47, 20);
		panel_1.add(textField_4);
		
		JLabel lblMailAutomationSoftware = new JLabel("MAIL AUTOMATION SOFTWARE");
		lblMailAutomationSoftware.setHorizontalAlignment(SwingConstants.CENTER);
		lblMailAutomationSoftware.setFont(new Font("Times New Roman", Font.BOLD, 16));
		lblMailAutomationSoftware.setForeground(Color.YELLOW);
		lblMailAutomationSoftware.setBounds(255, 12, 508, 14);
		frame.getContentPane().add(lblMailAutomationSoftware);
		
		textField_5 = new JTextField();
		textField_5.setHorizontalAlignment(SwingConstants.CENTER);
		textField_5.setEditable(false);
		textField_5.setColumns(10);
		textField_5.setBounds(284, 235, 47, 20);
		panel_1.add(textField_5);
		
		label = new JLabel("");
		label.setOpaque(true);
		label.setHorizontalAlignment(SwingConstants.CENTER);
		label.setForeground(Color.YELLOW);
		label.setFont(new Font("Times New Roman", Font.BOLD, 15));
		label.setBackground(Color.BLACK);
		label.setBounds(0, 347, 341, 21);
		panel_1.add(label);
		
		JLabel lblTotalUnreadMessages = new JLabel("TOTAL UNREAD MESSAGES REPLIED:");
		lblTotalUnreadMessages.setOpaque(true);
		lblTotalUnreadMessages.setHorizontalAlignment(SwingConstants.CENTER);
		lblTotalUnreadMessages.setForeground(Color.YELLOW);
		lblTotalUnreadMessages.setFont(new Font("Times New Roman", Font.BOLD, 11));
		lblTotalUnreadMessages.setBackground(Color.BLACK);
		lblTotalUnreadMessages.setBounds(0, 288, 258, 21);
		panel_1.add(lblTotalUnreadMessages);
		
		textField_6 = new JTextField();
		textField_6.setHorizontalAlignment(SwingConstants.CENTER);
		textField_6.setEditable(false);
		textField_6.setColumns(10);
		textField_6.setBounds(284, 288, 47, 20);
		panel_1.add(textField_6);
		
		JLabel lblDashboard = new JLabel("DASHBOARD");
		lblDashboard.setFont(new Font("Times New Roman", Font.BOLD, 14));
		lblDashboard.setHorizontalAlignment(SwingConstants.CENTER);
		lblDashboard.setForeground(Color.WHITE);
		lblDashboard.setBackground(Color.BLACK);
		lblDashboard.setBounds(862, 68, 155, 20);
		frame.getContentPane().add(lblDashboard);
		
		JLabel lblNewLabel = new JLabel("SUCCESSFULL LOGIN:");
		lblNewLabel.setBackground(Color.BLACK);
		lblNewLabel.setFont(new Font("Times New Roman", Font.BOLD, 13));
		lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel.setBounds(0, 0, 46, 14);
		frame.getContentPane().add(lblNewLabel);
		
		table = new JTable();
		table.setSurrendersFocusOnKeystroke(true);
		table.setForeground(UIManager.getColor("CheckBox.focus"));
		table.setBackground(UIManager.getColor("Button.foreground"));
		table.setFont(new Font("Times New Roman", Font.BOLD, 14));
		table.setModel(new DefaultTableModel(
			new Object[][] {
			},
			new String[] {
				"ID", "USER", "PASSWORD", "PROXY", "STATUS"
			}
		) {
			boolean[] columnEditables = new boolean[] {
				true, false, false, false, true
			};
			public boolean isCellEditable(int row, int column) {
				return columnEditables[column];
			}
		});
		table.getColumnModel().getColumn(1).setResizable(false);
		table.getColumnModel().getColumn(2).setResizable(false);
		table.getColumnModel().getColumn(3).setResizable(false);
		table.getColumnModel().getColumn(4).setPreferredWidth(392);
		table.getColumnModel().getColumn(4).setMinWidth(187);
		scrollPane.setViewportView(table);
		
ImageIcon img =new ImageIcon(getClass().getResource("images/mail.jpg"));
        
       
        
        frame.setIconImage(img.getImage());
		
	//	table.getColumnModel().getColumn(4).
		
		
		  table.addMouseListener(new MouseAdapter() {
			    public void mousePressed(MouseEvent mouseEvent) {
			        JTable table =(JTable) mouseEvent.getSource();
			        int row=table.getSelectedRow();
			       try{
			        System.out.println("row is : "+row);
			        
			        
			           
			       
			         
			            
			            String table_user=(table.getModel().getValueAt(row, 1)).toString();
				           
			            System.out.println("user: "+table_user);
			            
			            String table_pass=(table.getModel().getValueAt(row, 2)).toString();
				           
			            System.out.println("pass: "+table_pass);
			            
			            try{
			         table_proxy=(table.getModel().getValueAt(row, 3)).toString();
				           
			            System.out.println("proxy: "+table_proxy);
			            }catch(Exception df)
			            {
			            	
			            }
			            /*
			            String table_link=(table.getModel().getValueAt(row, 4)).toString();
				           
			            System.out.println("site link: "+table_link);
			            
			            String table_cvv=(table.getModel().getValueAt(row, 3)).toString();
				           
			            System.out.println("cvv: "+table_cvv);
			            
			            String head_less=(table.getModel().getValueAt(row, 6)).toString();
				           
			            System.out.println("cvv: "+head_less);
			            */
			            if (mouseEvent.getClickCount() == 2 && table.getSelectedRow() != -1) {
				            
				            // your valueChanged overridden method 
				        	  int a=JOptionPane.showConfirmDialog(frame,"Id: "+(row+1)+  "\nUsername: " +table_user+ "\n PassWord: "+table_pass+ "\n Proxy: "+table_proxy+ " \n Execute this task?" ); 
				        	  if(a==JOptionPane.YES_OPTION){  
				        		   
				        //	  nike_checkout(table_user,table_pass,table_proxy,table_cvv,table_link,row,head_less);
				        	//	  nike_checkout(String user,String pass,String proxy1,String link,String cvv,int r)
				        		  
				        		//  nike_checkout(table_user,table_pass,"0",table_link);
				        		  bet((row+1) ,table_user ,table_pass,table_proxy);
				        	  
				        }
				      
				        	  
				        	 /* if(a==JOptionPane.CANCEL_OPTION){  
				        		 
						           
						            System.out.println("task no: "+table_i);

									 Connection myConn =null;
								        Statement myStmt23=null;
								        Statement myStmt2=null;
								        ResultSet myRs23=null;
								         ResultSet myRs2 =null;
								        String dburl="jdbc:mysql://localhost:3306/nike_shoe";
								        String user="root";
								        String pass="";
								      
								      try{
								           
								          
								         //  String role1s= role.getText();
								           myConn = (Connection) DriverManager.getConnection(dburl, user, pass);
								           myStmt23=(Statement) myConn.createStatement();          
								        String sql23 = "DELETE FROM nike WHERE idNo='"+table_i+"'"; 
								        int up2 = myStmt23.executeUpdate(sql23);
							             if(up2>0)
							             {
							              //   System.out.println("successfully inserted data");
							                 
							                 JOptionPane.showMessageDialog(null,"TASK DELETED ..REFRESH THE TASK PAGE");
							                 
							            
							             }
								      }catch(Exception gf)
								      {
								    	  gf.printStackTrace();
								    	//  JOptionPane.showMessageDialog(null,"DATABASE CONNECTION ERROR");
								      }
					        		  
					        		} */
				        }
			       }catch(Exception df)
			       {
			    	   df.printStackTrace();
			       }
			      
			    }
			});
			
			 spinner = new JSpinner();
			 spinner.setForeground(Color.YELLOW);
			 spinner.setBackground(Color.YELLOW);
			spinner.setBounds(25, 152, 61, 20);
		//	frame.getContentPane().add(spinner);
			
		 spinner_1 = new JSpinner();
		 spinner_1.setForeground(Color.YELLOW);
		 spinner_1.setBackground(Color.YELLOW);
			spinner_1.setBounds(111, 152, 61, 20);
		//	frame.getContentPane().add(spinner_1);
			
			label_2 = new JLabel("");
			label_2.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseClicked(MouseEvent e) {
					
				//	String d=	textField.getText();
					/*
					if(d.isEmpty())
					{
						JOptionPane.showMessageDialog(null, "delay is empty");
					}
					*/
					/*
					else{
					
					int s = Integer.parseInt(d);
						
					 delay = s * 1000;
						
						
						System.out.println("delay in seconds is: "+delay);
						
						
						
						*/
						
						
						
						
						
						
						
						
			//			ADD_TASK add = new ADD_TASK();
						
				//		add.frame.setVisible(true);
				//		 model2.addRow(new Object[]{"asd", "asff", "sasaeg"});
						File excelFile;
				        FileInputStream excelFIS = null;
				        BufferedInputStream excelBIS = null;
				        XSSFWorkbook excelImportToJTable = null;
				        String defaultCurrentDirectoryPath = "C:\\Users\\Authentic\\Desktop";
				        JFileChooser excelFileChooser = new JFileChooser(defaultCurrentDirectoryPath);
				        excelFileChooser.setDialogTitle("Select Excel File");
				        FileNameExtensionFilter fnef = new FileNameExtensionFilter("EXCEL FILES", "xls", "xlsx", "xlsm", "csv");
				        excelFileChooser.setFileFilter(fnef);
				        int excelChooser = excelFileChooser.showOpenDialog(null);
				        if (excelChooser == JFileChooser.APPROVE_OPTION) {
				            try {
				                excelFile = excelFileChooser.getSelectedFile();
				                excelFIS = new FileInputStream(excelFile);
				                excelBIS = new BufferedInputStream(excelFIS);
				                excelImportToJTable = new XSSFWorkbook(excelBIS);
				                XSSFSheet excelSheet = excelImportToJTable.getSheetAt(0);
				 
				                for (int row = 1; row<= excelSheet.getLastRowNum(); row++) {
				                	
				                	
				                    XSSFRow excelRow = excelSheet.getRow(row);
				 
				                    try{
				                   ID = excelRow.getCell(0).getNumericCellValue();
				                    }catch(Exception fh)
				                    {
				                 
				                    }
				                   
				                   int id_in =(int)ID;
				                   
				                    
				                    
				                    XSSFCell USER = excelRow.getCell(1);
				                    
				                    if(USER == null)
				                	{
				                    	USER = excelRow.createCell(1);
				                	}
				                    
				                    XSSFCell PASS = excelRow.getCell(2);
				                    if(PASS == null)
				                	{
				                    	PASS = excelRow.createCell(2);
				                	}
				                    
				                    XSSFCell PROXY = excelRow.getCell(3);
				                    if(PROXY == null)
				                	{
				                    	PROXY = excelRow.createCell(3);
				                	}
				                    /*
				                   double   excelSubject = excelRow.getCell(3).getNumericCellValue();
				                    XSSFCell excelSubject2 = excelRow.getCell(4);
				                    XSSFCell HEAD = excelRow.getCell(5);
				                    
				                    
				                   int cv =(int)excelSubject;
				                //    XSSFCell excelImage = excelRow.getCell(4);
				               //     model2.addRow(new Object[]{name3,bre ,luna,dina,cvv,tota}
				             //       JLabel excelJL = new JLabel(new ImageIcon(new ImageIcon(excelImage.getStringCellValue()).getImage().getScaledInstance(60, 60, Image.SCALE_SMOOTH)));
				                 */
				                    model2 = (DefaultTableModel) table.getModel();
				                    model2.addRow(new Object[]{id_in, USER, PASS,PROXY});
				                    
				                
				                }
				                imported = true;
				                path = excelFile.getAbsolutePath();
				                JOptionPane.showMessageDialog(null, "Import was Successfull !!.....");
				            } catch (IOException iOException) {
				                JOptionPane.showMessageDialog(null, iOException.getMessage());
				            } finally {
				                try {
				                    if (excelFIS != null) {
				                        excelFIS.close();
				                    }
				                    if (excelBIS != null) {
				                        excelBIS.close();
				                    }
				                    if (excelImportToJTable != null) {
				                        excelImportToJTable.close();
				                    }
				                } catch (IOException iOException) {
				                    JOptionPane.showMessageDialog(null, iOException.getMessage());
				                }
				            }
				        }
					}
						
					
					
				
				@Override
				public void mouseEntered(MouseEvent e) {
					label_2.setIcon(new ImageIcon(getClass().getResource("images/button (46).png")));
					
				}
				@Override
				public void mouseExited(MouseEvent e) {
					label_2.setIcon(new ImageIcon(getClass().getResource("images/button (45).png")));
				}
			});
			
			label_2.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			label_2.setHorizontalAlignment(SwingConstants.CENTER);
			label_2.setIcon(new ImageIcon(getClass().getResource("images/button (45).png")));
			label_2.setBounds(25, 68, 147, 48);
			frame.getContentPane().add(label_2);
			
			label_1 = new JLabel("IMPORT EXCEL");
			label_1.setHorizontalAlignment(SwingConstants.CENTER);
			label_1.setForeground(Color.YELLOW);
			label_1.setFont(new Font("Times New Roman", Font.BOLD, 15));
			label_1.setBounds(25, 43, 147, 14);
			frame.getContentPane().add(label_1);
			
			label_3 = new JLabel("");
			label_3.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			label_3.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseClicked(MouseEvent e) {
					
					 hour_j = (Integer)spinner.getValue();
		        		
	        		 min_j = (Integer)spinner_1.getValue();
	        		 
	        		 if(textField_1.getText().isEmpty() || textField_2.getText().isEmpty() || path.isEmpty())
	        		 {
	        			 JOptionPane.showMessageDialog(null,"Please insert range / Import excel sheet");
	        		 }else{
	        			// bet5();
	        			 String value1 = comboBox.getSelectedItem().toString();
	        		 JOptionPane.showMessageDialog(null,"All tasks will be repeated in "+value1+" hour");
	        		 }
	        		 
				}
				@Override
				public void mouseEntered(MouseEvent e) {
					label_3.setIcon(new ImageIcon(getClass().getResource("images/button (52).png")));
				}
				@Override
				public void mouseExited(MouseEvent e) {
					label_3.setIcon(new ImageIcon(getClass().getResource("images/button (51).png")));
				}
			});
			
			lblScheduleTasks = new JLabel("SCHEDULE TASKS");
			lblScheduleTasks.setForeground(Color.YELLOW);
			lblScheduleTasks.setFont(new Font("Times New Roman", Font.BOLD, 13));
			lblScheduleTasks.setHorizontalAlignment(SwingConstants.CENTER);
			lblScheduleTasks.setBounds(29, 127, 139, 14);
			frame.getContentPane().add(lblScheduleTasks);
			
			lblClickExecute = new JLabel("1 CLICK EXECUTE ");
			lblClickExecute.setForeground(Color.YELLOW);
			lblClickExecute.setFont(new Font("Times New Roman", Font.BOLD, 14));
			lblClickExecute.setHorizontalAlignment(SwingConstants.CENTER);
			lblClickExecute.setBounds(25, 385, 143, 15);
			frame.getContentPane().add(lblClickExecute);
			label_3.setHorizontalAlignment(SwingConstants.CENTER);
			label_3.setIcon(new ImageIcon(getClass().getResource("images/button (51).png")));
			
			label_3.setBounds(10, 183, 172, 50);
			frame.getContentPane().add(label_3);
			
			label_4 = new JLabel("");
			label_4.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseEntered(MouseEvent e) {
					label_4.setIcon(new ImageIcon(getClass().getResource("images/button (56).png")));
				}
				@Override
				public void mouseExited(MouseEvent e) {
					label_4.setIcon(new ImageIcon(getClass().getResource("images/button (55).png")));
				}
				@Override
				public void mouseClicked(MouseEvent e) {
					 String value1 = comboBox.getSelectedItem().toString();
					 
					 if(value1=="SELECT HOUR")
					 {
						 JOptionPane.showMessageDialog(null,"Please select waiting hour");
					 }else{
					if(imported == true && value1!="SELECT HOUR")
					{
						
						 s_l=0;
						 f_l=0;
						 t_s=0;
					     t_r=0;
						nike_3();
						JOptionPane.showMessageDialog(null,"All task have been executed");
						
					}else{
						JOptionPane.showMessageDialog(null, "You haven't imported any excel sheet/tasks");
					}
					 }
					
					
				}
			});
			label_4.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			label_4.setHorizontalAlignment(SwingConstants.CENTER);
			label_4.setBounds(29, 408, 143, 39);
			
			label_4.setIcon(new ImageIcon(getClass().getResource("images/button (55).png")));
			frame.getContentPane().add(label_4);
			
			JLabel lblV = new JLabel("V 2.2");
			lblV.setFont(new Font("Times New Roman", Font.BOLD, 20));
			lblV.setForeground(Color.YELLOW);
			lblV.setHorizontalAlignment(SwingConstants.LEFT);
			lblV.setBounds(10, 11, 92, 20);
			frame.getContentPane().add(lblV);
			
			lblSelectSpecificTask = new JLabel("SELECT RANGE");
			lblSelectSpecificTask.setHorizontalAlignment(SwingConstants.CENTER);
			lblSelectSpecificTask.setFont(new Font("Times New Roman", Font.BOLD, 13));
			lblSelectSpecificTask.setForeground(Color.YELLOW);
			lblSelectSpecificTask.setBounds(25, 239, 147, 14);
			frame.getContentPane().add(lblSelectSpecificTask);
			
			textField_1 = new JTextField();
			textField_1.setBounds(25, 264, 37, 20);
			frame.getContentPane().add(textField_1);
			textField_1.setColumns(10);
			
			textField_2 = new JTextField();
			textField_2.setBounds(135, 264, 37, 20);
			frame.getContentPane().add(textField_2);
			textField_2.setColumns(10);
			
			JLabel label_5 = new JLabel("");
			label_5.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseClicked(MouseEvent e) {
			
					 s_l=0;
					 f_l=0;
					 t_s=0;
				     t_r=0;
				     t_u_r=0;
				     String value1 = comboBox.getSelectedItem().toString();
				     
				     if(value1=="SELECT HOUR")
				     {
					     
				    	 JOptionPane.showMessageDialog(null, "Please select waiting hour");
				      }else{
				    	  nike_2();
						     JOptionPane.showMessageDialog(null, "Ranges have been executed");
				      }
				
				}
			});
			label_5.setHorizontalAlignment(SwingConstants.CENTER);
			label_5.setBounds(25, 329, 143, 45);
			
			label_5.setIcon(new ImageIcon(getClass().getResource("images/button (57).png")));
			
			label_5.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			frame.getContentPane().add(label_5);
			
			lblTo = new JLabel("TO");
			lblTo.setFont(new Font("Times New Roman", Font.BOLD, 14));
			lblTo.setHorizontalAlignment(SwingConstants.CENTER);
			lblTo.setForeground(Color.YELLOW);
			lblTo.setBounds(68, 264, 46, 20);
			frame.getContentPane().add(lblTo);
			
			comboBox = new JComboBox();
			comboBox.setBackground(Color.YELLOW);
			comboBox.setModel(new DefaultComboBoxModel(new String[] {"SELECT HOUR", "1", "2", "4", "6", "8", "10", "12", "15", "24"}));
			comboBox.setBounds(25, 152, 143, 20);
			frame.getContentPane().add(comboBox);
	}
	public static void bet(final int id, final String table_user,final String table_pass, final String table_proxy2){
		SwingWorker<Void,Void> email = new SwingWorker<Void,Void>()
				{

					@Override
					protected Void doInBackground() throws Exception {
						
						
						for(int y =0;y<10;y++)
						 {
							DayOfWeek dayOfWeek2 = DayOfWeek.from(LocalDate.now());
							
			        		
			        		String d = String.valueOf(dayOfWeek2);
			        	//	System.out.println(d);
			        		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH:mm");  
				        	   LocalDateTime now = LocalDateTime.now(); 
				        	   String da=dtf.format(now);
				        	 //  System.out.println(da);  
				        	   
				        	  // textField_5.setText(d+" "+da);
						
				        	   System.out.println(d+" "+da);
				        	   
				        	   String [] ho = da.split(":");
				        	   String h =ho[0];
				        	   String m =ho[1];
							
							/*
							if((hour_j <= Integer.parseInt(h)) &&(min_j <= Integer.parseInt(m)))
								
							{
									
									y=200;
									model2.setValueAt("time is up!...starting the task", (id-1),4 );
									System.out.println("scheduled time is up,execute row :" +id);
									
									*/
									
						 try{
							 IMAPFolder folder = null;
							 
							 IMAPFolder folder_spam = null;
							 
						        Store store = null;
						        String subject = null;
						        Flag flag = null;
						        try 
						        {
						            
						        
						        	
						        	
						       //   Properties props = 
						          String proxyIP = "proxy.packetstream.io";
						            String proxyPort = "31112";
						            String proxyUser = "stone";
						            String proxyPassword = "O0R7TU9ODuBgePdO";
						            prop = System.getProperties();
						          //  prop.put("proxySet", "true");
						          // prop.setProperty("mail.imaps.proxy.host", proxyIP);
						        //    prop.setProperty("mail.imaps.proxy.port", proxyPort);
						          //  prop.setProperty("mail.imaps.proxy.user", proxyUser);
						        //    prop.setProperty("mail.imaps.proxy.password", proxyPassword);
						          prop.setProperty("mail.store.protocol", "imaps");
						            
						          //  prop.setProperty("proxySet","true"); 
						            
						      	try{
									String proxy_u =table_proxy2;
									String[] p = proxy_u.split (":");
								 ip=p[0].trim();
								
								 port=p[1].trim();
								
								 prop.setProperty("socksProxyHost",ip); 
						            
						            prop.setProperty("socksProxyPort",port);
								
									}catch(Exception ds)
									{
										
									}  
						          
						          
						          
						       int s_l_l=0;
						       int t_s_l=0;
						       int f_l_l=0;
						       int t_r_l=0;
						       int t_u_r_l=0;

						          Session session = Session.getDefaultInstance(prop, null);
						 
						          

						          store = session.getStore("imaps");

							    	Random r = new Random();
									 int low = 1000;
									 int high = 5000;
									 
									 int ran   = r.nextInt(high-low) + low;
									 
									 
						          
						        //  store.connect("imap.mail.yahoo.com","milan_halvorson90@yahoo.com", "duflbqleqbdyvpln");
						        
						          model2.setValueAt("loging into the account", (id-1),4 );
									System.out.println("loging into the account " +(id-1));
									
									
									
									if(table_user.contains("gmail.com"))
									{	
										
										try{
									
						          store.connect("imap.gmail.com",table_user, table_pass);
						        
						          model2.setValueAt("logged in successfully", (id-1),4 );
						         s_l_l=1;
						         
						         s_l = s_l+s_l_l;
									System.out.println("logged in successfully :" +(id-1));
									
									  low = 1000;
									  high = 5000;
									 
									  ran   = r.nextInt(high-low) + low;
									 Thread.sleep(ran);
						          // store.connect("imap.aol.com","shemikamoloney653237@gmail.com", "lkpxkklueo");
						          folder = (IMAPFolder) store.getFolder("inbox"); // This doesn't work for other email account
						          folder_spam = (IMAPFolder) store.getFolder("[Gmail]/Spam");
						          Folder[] f = store.getDefaultFolder().list();
						        		  
						          for(Folder fd:f)
						              System.out.println(">> "+fd.getName());
						          
						          //folder = (IMAPFolder) store.getFolder("inbox"); This works for both email account
						          model2.setValueAt("Unmarking spam messages as not spam", (id-1),4 );
						          if(!folder_spam.isOpen())
							          folder_spam.open(Folder.READ_WRITE);
							          Message[] messages_spam = folder_spam.getMessages();
						          
							          System.out.println("No of Spam Messages found: " + folder_spam.getMessageCount());
							          model2.setValueAt("No of Spam Messages found : " + folder_spam.getMessageCount(), (id-1),4 );
							          t_s_l = folder_spam.getMessageCount();
							          t_s=t_s+t_s_l;
						          int sp = folder_spam.getMessageCount();
						          if (sp != 0) {
						        	  try{
						              folder_spam.copyMessages(messages_spam, folder);
						              folder_spam.setFlags(messages_spam, new Flags(Flags.Flag.DELETED), true);
						        	  }catch(Exception tr)
						        	  {
						        		  
						        	  }
						        	  

									  low = 1000;
									  high = 5000;
									 
									  ran   = r.nextInt(high-low) + low;
									 Thread.sleep(ran);
						              // Dump out the Flags of the moved messages, to insure that
						              // all got deleted
						        	  try{
						              for (int i = 0; i < sp; i++) {
						                if (!messages_spam[i].isSet(Flags.Flag.DELETED))
						                  System.out.println("Message # " + messages_spam[i] + " not deleted");
						              }
						        	  }catch(Exception rt)
						        	  {
						        		  
						        	  }
						        	  
						        	  model2.setValueAt("All spam messages have been marked as not spam!", (id-1),4 );
						            }
						          
						       
						          if(!folder.isOpen())
						          folder.open(Folder.READ_WRITE);
						          Message[] messages = folder.getMessages();
						        //  System.out.println("No of Messages : " + folder.getMessageCount());
						          System.out.println("No of Unread Messages : " + folder.getUnreadMessageCount());
						          model2.setValueAt("No of Unread Messages : " + folder.getUnreadMessageCount(), (id-1),4 );
						          
						          t_r_l = folder.getUnreadMessageCount();
						          t_r = t_r +  t_r_l ;
						          System.out.println(messages.length);
						          
						       // Fetch unseen messages from inbox folder
						          
						          model2.setValueAt("Getting unread emails from inbox", (id-1),4 );
						          
						          Message[] messages2 = folder.search(
						              new FlagTerm(new Flags(Flags.Flag.SEEN), false));
						          /*
						          Arrays.sort( messages2, ( m1, m2 ) -> {
						              try {
						                return m2.getSentDate().compareTo( m1.getSentDate() );
						              } catch ( MessagingException e ) {
						                throw new RuntimeException( e );
						              }
						            } );
						          */
						          model2.setValueAt("Reading  all unread emails from inbox", (id-1),4 );
						          int un = folder.getUnreadMessageCount();
						          for (int i=0; i < un;i++) 
						          {
						        	  model2.setValueAt("Reading   unread email  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 );

						            System.out.println("*****************************************************************************");
						            System.out.println("MESSAGE " + (i + 1) + ":");
						            //reads only unseen messages
						            Message msg =  messages2[i];
						            
						            
						            //reads all the messages in inbox
						       //    Message msg =  messages[i];
						            
						            //System.out.println(msg.getMessageNumber());
						            //Object String;
						            //System.out.println(folder.getUID(msg)

						            subject = msg.getSubject();
						            
						          

						            System.out.println("Subject: " + subject);
						            System.out.println("From: " + msg.getFrom()[0]);
						           System.out.println("To: "+msg.getAllRecipients()[0]);
						            System.out.println("Date: "+msg.getReceivedDate());
						            System.out.println("Size: "+msg.getSize());
						            System.out.println(msg.getFlags());
						            System.out.println("Body: \n"+ msg.getContent());
						            System.out.println(msg.getContentType());
						            
						          //  msg.setFlag(Flags.Flag.DELETED, true);
						            
						        //    msg.getFolder().copyMessages(messages2, folder_spam);
						            
						        //    folder.copyMessages(messages, folder_spam); //Moves email msg to SPAM folder
						            
						            msg.setFlag(Flags.Flag.SEEN, true);
						            
						            
						            if(Reply == true)
								          

							          {
						            
						            Date     date = msg.getSentDate();
						               // Get all the information from the message
						               String from = InternetAddress.toString(msg.getFrom());
						               if (from != null) {
						                  System.out.println("From: " + from);
						               }
						               String replyTo = InternetAddress.toString(msg
							         .getReplyTo());
						               if (replyTo != null) {
						                  System.out.println("Reply-to: " + replyTo);
						               }
						               String to = InternetAddress.toString(msg
							         .getRecipients(Message.RecipientType.TO));
						               if (to != null) {
						                  System.out.println("To: " + to);
						               }

						               String subject2 = msg.getSubject();
						               if (subject2 != null) {
						                  System.out.println("Subject: " + subject2);
						               }
						               Date sent = msg.getSentDate();
						               if (sent != null) {
						                  System.out.println("Sent: " + sent);
						               }
						               
						               
						               
						               System.out.println("\n 1st ===> setup Mail Server Properties..");
										mailServerProperties = System.getProperties();
										mailServerProperties.put("mail.smtp.port", "587");
										mailServerProperties.put("mail.smtp.auth", "true");
										mailServerProperties.put("mail.smtp.starttls.enable", "true");
										
								     	try{
											String proxy_u =table_proxy2;
											String[] p = proxy_u.split (":");
										 ip=p[0].trim();
										
										 port=p[1].trim();
										
										 mailServerProperties.setProperty("socksProxyHost",ip); 
								            
										 mailServerProperties.setProperty("socksProxyPort",port);
										
											}catch(Exception ds)
											{
												
											} 
								     	
								        InputStream inp = new FileInputStream(path);
										 Workbook wb = WorkbookFactory.create(inp);
										 Sheet sheet = wb.getSheetAt(0);
										 int ctr  = 1;
										 Row row = null;
									      Cell cell =null; 
									      
									      row = sheet.getRow(id);
									      
									      cell = row.getCell(4);
									      
									      if(cell == null)
									      {
									    	  cell = row.createCell(4);
									      }
										System.out.println("Mail Server Properties have been setup successfully..");
										getMailSession = Session.getDefaultInstance(mailServerProperties, null);
										
										Message replyMessage = new MimeMessage(getMailSession);
						                  replyMessage = (MimeMessage) msg.reply(false);
						                  replyMessage.setFrom(new InternetAddress(to));
						                  replyMessage.setText(cell.toString());
						                  replyMessage.setReplyTo(msg.getReplyTo());

						                  // Send the message by authenticating the SMTP server
						                  // Create a Transport instance and call the sendMessage
						                  Transport t = getMailSession.getTransport("smtp");
						                  
						                  try {
							   	     //connect to the smpt server using transport instance
								     //change the user and password accordingly	
							             t.connect("smtp.gmail.com",table_user, table_pass);
							             t.sendMessage(replyMessage,
						                        replyMessage.getAllRecipients());
						                  } finally {
						                     t.close();
						                  }
						                  System.out.println("message replied successfully ....");
						                  
						                  ++t_u_r_l;
						                  t_u_r=t_u_r+t_u_r_l;
						                  
						               


						            
						            model2.setValueAt("Unread email read and replied successfully  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 );

						          }else{
						        	  model2.setValueAt("Unread email read successfully  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 );
						          }
						          model2.setValueAt("All unread emails have been read and replied", (id-1),4 );
										}
						          model2.setValueAt("Tasks completed for this account", (id-1),4 );
										}catch(Exception fh)
										{
											fh.printStackTrace();
											 model2.setValueAt("Operation failed due to incorrect login details", (id-1),4 );
											 
											 f_l_l=1;
											 f_l = f_l+f_l_l;
										}
							          
							          
									}
									if(table_user.contains("yahoo.com"))
									{	
									try{
						          store.connect("imap.mail.yahoo.com",table_user, table_pass);
						        
						          model2.setValueAt("logged in successfully", (id-1),4 );
									System.out.println("logged in successfully :" +(id-1));
									
									 s_l_l=1;
							         
							         s_l = s_l+s_l_l;
						          // store.connect("imap.aol.com","shemikamoloney653237@gmail.com", "lkpxkklueo");
							         

									  low = 1000;
									  high = 5000;
									 
									  ran   = r.nextInt(high-low) + low;
									 Thread.sleep(ran);
						          folder = (IMAPFolder) store.getFolder("Inbox"); // This doesn't work for other email account
						          folder_spam = (IMAPFolder) store.getFolder("Bulk Mail");
						          Folder[] f = store.getDefaultFolder().list();
						        		  
						          for(Folder fd:f)
						              System.out.println(">> "+fd.getName());
						          
						          //folder = (IMAPFolder) store.getFolder("inbox"); This works for both email account
						          model2.setValueAt("Unmarking spam messages as not spam", (id-1),4 );
						          if(!folder_spam.isOpen())
							          folder_spam.open(Folder.READ_WRITE);
							          Message[] messages_spam = folder_spam.getMessages();
						          
							          System.out.println("No of Spam Messages found: " + folder_spam.getMessageCount());
							          model2.setValueAt("No of Spam Messages found : " + folder_spam.getMessageCount(), (id-1),4 );
							          
							          t_s_l = folder_spam.getMessageCount();
							          t_s=t_s+t_s_l;
						          int sp = folder_spam.getMessageCount();
						          if (sp != 0) {
						        	  try{
						              folder_spam.copyMessages(messages_spam, folder);
						              folder_spam.setFlags(messages_spam, new Flags(Flags.Flag.DELETED), true);
						        	  }catch(Exception tr)
						        	  {
						        		  
						        	  }
						        	  

									  low = 1000;
									  high = 5000;
									 
									  ran   = r.nextInt(high-low) + low;
									 Thread.sleep(ran);
						              // Dump out the Flags of the moved messages, to insure that
						              // all got deleted
						        	  try{
						              for (int i = 0; i < sp; i++) {
						                if (!messages_spam[i].isSet(Flags.Flag.DELETED))
						                  System.out.println("Message # " + messages_spam[i] + " not deleted");
						              }
						        	  }catch(Exception rt)
						        	  {
						        		  
						        	  }
						        	  
						        	  model2.setValueAt("All spam messages have been marked as not spam!", (id-1),4 );
						            }
						        

						          if(!folder.isOpen())
						          folder.open(Folder.READ_WRITE);
						          Message[] messages = folder.getMessages();
						        //  System.out.println("No of Messages : " + folder.getMessageCount());
						          System.out.println("No of Unread Messages : " + folder.getUnreadMessageCount());
						          model2.setValueAt("No of Unread Messages : " + folder.getUnreadMessageCount(), (id-1),4 );
						          System.out.println(messages.length);
						          
						          t_r_l = folder.getUnreadMessageCount();
						          t_r = t_r +  t_r_l ;
						          
						       // Fetch unseen messages from inbox folder
						          
						          model2.setValueAt("Getting unread emails from inbox", (id-1),4 );
						          
						          Message[] messages2 = folder.search(
						              new FlagTerm(new Flags(Flags.Flag.SEEN), false));
						          /*
						          Arrays.sort( messages2, ( m1, m2 ) -> {
						              try {
						                return m2.getSentDate().compareTo( m1.getSentDate() );
						              } catch ( MessagingException e ) {
						                throw new RuntimeException( e );
						              }
						            } );
						            */
						          model2.setValueAt("Reading  all unread emails from inbox", (id-1),4 );
						          int un = folder.getUnreadMessageCount();
						          for (int i=0; i < un;i++) 
						          {
						        	  model2.setValueAt("Reading   unread email  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 );

						            System.out.println("*****************************************************************************");
						            System.out.println("MESSAGE " + (i + 1) + ":");
						            //reads only unseen messages
						            Message msg =  messages2[i];
						            
						            
						            //reads all the messages in inbox
						       //    Message msg =  messages[i];
						            
						            //System.out.println(msg.getMessageNumber());
						            //Object String;
						            //System.out.println(folder.getUID(msg)

						            subject = msg.getSubject();
						            
						          

						            System.out.println("Subject: " + subject);
						            System.out.println("From: " + msg.getFrom()[0]);
						           System.out.println("To: "+msg.getAllRecipients()[0]);
						            System.out.println("Date: "+msg.getReceivedDate());
						            System.out.println("Size: "+msg.getSize());
						            System.out.println(msg.getFlags());
						            System.out.println("Body: \n"+ msg.getContent());
						            System.out.println(msg.getContentType());
						            
						          //  msg.setFlag(Flags.Flag.DELETED, true);
						            
						        //    msg.getFolder().copyMessages(messages2, folder_spam);
						            
						        //    folder.copyMessages(messages, folder_spam); //Moves email msg to SPAM folder
						            
						            msg.setFlag(Flags.Flag.SEEN, true);
						            if(Reply == true)
							        	  
							          {
						            Date     date = msg.getSentDate();
						               // Get all the information from the message
						               String from = InternetAddress.toString(msg.getFrom());
						               if (from != null) {
						                  System.out.println("From: " + from);
						               }
						               String replyTo = InternetAddress.toString(msg
							         .getReplyTo());
						               if (replyTo != null) {
						                  System.out.println("Reply-to: " + replyTo);
						               }
						               String to = InternetAddress.toString(msg
							         .getRecipients(Message.RecipientType.TO));
						               if (to != null) {
						                  System.out.println("To: " + to);
						               }

						               String subject2 = msg.getSubject();
						               if (subject2 != null) {
						                  System.out.println("Subject: " + subject2);
						               }
						               Date sent = msg.getSentDate();
						               if (sent != null) {
						                  System.out.println("Sent: " + sent);
						               }
						               
						               
						               InputStream inp = new FileInputStream(path);
										 Workbook wb = WorkbookFactory.create(inp);
										 Sheet sheet = wb.getSheetAt(0);
										 int ctr  = 1;
										 Row row = null;
									      Cell cell =null; 
									      
									      row = sheet.getRow(id);
									      
									      cell = row.getCell(4);
									      
									      if(cell == null)
									      {
									    	  cell = row.createCell(4);
									      }
										
						               
						               String from2=table_user;//change accordingly
						               String password=table_pass;//change accordingly

						               //Get the session object
						               /*
						               Properties props = new Properties();
						               props.put("mail.smtp.host", "smtp.mail.yahoo.com");
						               props.put("mail.smtp.socketFactory.port", "465");
						               props.put("mail.smtp.socketFactory.class",
						               "javax.net.ssl.SSLSocketFactory");
						               props.put("mail.smtp.auth", "true");
						               props.put("mail.smtp.port", "465");
						               props.put("mail.debug", "true");
						               props.put("mail.store.protocol", "imap");
						               props.put("mail.transport.protocol", "smtp");
						               
						               props.put("mail.smtp.starttls.enable", "true");
						               
						               try{
											String proxy_u =table_proxy2;
											String[] p = proxy_u.split (":");
										 ip=p[0].trim();
										
										 port=p[1].trim();
										
										 props.setProperty("socksProxyHost",ip); 
								            
										 props.setProperty("socksProxyPort",port);
										
											}catch(Exception ds)
											{
												
											} 
						               
						               Session session2 = Session.getInstance(props,
						               new javax.mail.Authenticator() {
						               protected PasswordAuthentication getPasswordAuthentication() {
						               return new PasswordAuthentication(from2,password);
						               }
						               });

						               //compose message
						          
						               MimeMessage message = new MimeMessage(session2);
						            //   message =  (MimeMessage) msg.reply(true);
						               message.setFrom(new InternetAddress(from2));
						               message.addRecipient(Message.RecipientType.TO,new InternetAddress(replyTo));
						               message.setSubject(subject2);
						               message.setText(cell.toString());
						               
						               
						               Transport.send(message);
						               System.out.println("message replied successfully ....");
						               */
						           
						               //send message
						               /*
						              MimeMessage replyMessage = new MimeMessage(session2);
						                  replyMessage = (MimeMessage) msg.reply(false);
						                  replyMessage.setFrom(new InternetAddress(from2));
						                  replyMessage.setText("Thanks");
						                  replyMessage.setReplyTo(msg.getReplyTo());
						               
						               */
						               /*
						               Message reply = msg.reply(false);
						               reply.addRecipient(Message.RecipientType.TO,new InternetAddress(to2));
						               reply.setText("THANKS");
						               
						               
						               Transport t = session.getTransport("smtp");
						                  try {
							   	     //connect to the smpt server using transport instance
								     //change the user and password accordingly	
							             t.connect("smtp.mail.yahoo.com",from2, password);
							             t.sendMessage(reply,
						                        reply.getAllRecipients());
						                  } finally {
						                     t.close();
						                  }
						                
						             */
						             //  String to2="mahbub.shaun@gmail.com";//change accordingly
						              // String from2="pwatkins208@yahoo.com";//change accordingly
						       //        String password="fgxbfovgigpfyldz";//change accordingly
						               
						                  System.out.println("\n 1st ===> setup Mail Server Properties..");
											mailServerProperties = System.getProperties();
											mailServerProperties.put("mail.smtp.host", "smtp.mail.yahoo.com");
											mailServerProperties.put("mail.smtp.socketFactory.port", "465");
											mailServerProperties.put("mail.smtp.socketFactory.class",
								               "javax.net.ssl.SSLSocketFactory");
											mailServerProperties.put("mail.smtp.auth", "true");
											mailServerProperties.put("mail.smtp.port", "465");
											mailServerProperties.put("mail.debug", "true");
											mailServerProperties.put("mail.store.protocol", "imap");
											mailServerProperties.put("mail.transport.protocol", "smtp");
								               
											mailServerProperties.put("mail.smtp.starttls.enable", "true");
										
								               
										//	mailServerProperties.put("mail.smtp.starttls.enable", "true");
											
									     	try{
												String proxy_u =table_proxy2;
												String[] p = proxy_u.split (":");
											 ip=p[0].trim();
											
											 port=p[1].trim();
											
											 mailServerProperties.setProperty("socksProxyHost",ip); 
									            
											 mailServerProperties.setProperty("socksProxyPort",port);
											
												}catch(Exception ds)
												{
													
												} 
											System.out.println("Mail Server Properties have been setup successfully..");
											getMailSession = Session.getDefaultInstance(mailServerProperties, null);
											
											/*
											Message replyMessage = new MimeMessage(getMailSession);
							                  replyMessage = (MimeMessage) msg.reply(false);
							                  replyMessage.setFrom(new InternetAddress(to));
							                  replyMessage.setText(cell.toString());
							              //    replyMessage.setReplyTo(msg.getReplyTo());
							                  
							                  replyMessage.addRecipient(Message.RecipientType.TO,new InternetAddress(replyTo));

							                  // Send the message by authenticating the SMTP server
							                  // Create a Transport instance and call the sendMessage
							                  Transport t = session.getTransport("smtp");
							                  try {
								   	     //connect to the smpt server using transport instance
									     //change the user and password accordingly	
								             t.connect("smtp.mail.yahoo.com",from2, password);
								             t.sendMessage(replyMessage,
							                        replyMessage.getAllRecipients());
							                  } finally {
							                     t.close();
							                  }
							                  */
											
											//set the addresses, to and from
											/*
											InternetAddress fromAddress;
											
											fromAddress = new InternetAddress("bettylittleton231@yahoo.com");
											msg.setFrom(fromAddress);

										
											//since mail can be sent to more than one recipient, create loop
											//to add all addresses into InternetAddress, addressTo.
											//InternetAddress[] toAddress = new InternetAddress[recipients.length];
										
											InternetAddress toAddress;
											
											toAddress = new InternetAddress("mdmahbubshaun@gmail.com");
											msg.setRecipient(Message.RecipientType.TO, toAddress);
											
											*/
											 String[] toa = {"mdmahbubshaun@gmail.com" };
											 InternetAddress[] toAddress = new InternetAddress[toa.length];
											
											toAddress[0]=new InternetAddress(toa[0]);

											MimeMessage reply = (MimeMessage)msg.reply(false);
											reply.setFrom(new InternetAddress(from2));
											reply.setText(cell.toString());
											  Session session2 = Session.getInstance(mailServerProperties,
										               new javax.mail.Authenticator() {
										               protected PasswordAuthentication getPasswordAuthentication() {
										               return new PasswordAuthentication(from2,password);
										               }
										               });

											Transport.send(reply,from2,password);
											
											   ++t_u_r_l;
								                  t_u_r=t_u_r+t_u_r_l;
							                  
							                  
											/*
											 String text = (String)msg.getContent();
											    Message reply = msg.
											    String replyText = text.replaceAll("(?m)^", "> ");
											    // allow user to edit replyText,
											    // e.g., using a Swing GUI or a web form
											    reply.setText(replyText);
											    reply.setReplyTo(msg.getReplyTo());
											    */
											    /*
											Message replyMessage = new MimeMessage(getMailSession);
							                  replyMessage = (MimeMessage) msg.reply(false);
							                  replyMessage.setFrom(new InternetAddress(to));
							                  replyMessage.setText(cell.toString());
							                  replyMessage.setReplyTo(msg.getReplyTo());
*/
							                  // Send the message by authenticating the SMTP server
							                  // Create a Transport instance and call the sendMessage
											/*
							                  Transport t = session.getTransport("smtp");
							                  try {
								   	     //connect to the smpt server using transport instance
									     //change the user and password accordingly	
								             t.connect("smtp.mail.yahoo.com",from2, password);
								             t.sendMessage(reply,
							                        reply.getAllRecipients());
							                  } finally {
							                     t.close();
							                  }
							                  */
							                  /*
											boolean attachment = false;
											 String host = "smtp.mail.yahoo.com";
											    String username = from2;
											    String pass = password;
											    Properties props = System.getProperties();
											    props.put("mail.smtp.starttls.enable", "true"); // added this line
											    props.put("mail.smtp.host", host);
											    props.put("mail.smtp.user", username);
											    props.put("mail.smtp.password", pass);
											    props.put("mail.smtp.port", "587");
											    props.put("mail.smtp.auth", "true");

											//    Session session = Session.getDefaultInstance(props);

											    MimeMessage mimeReply = new MimeMessage(getMailSession);
											    mimeReply.setFrom((Address) InternetAddress.parse(from)[0]);        

											    BodyPart messageBodyPart = new MimeBodyPart();
											    messageBodyPart.setContent(cell.toString(), "text/html");

											    Multipart multipart = new MimeMultipart();
											    // Set text message part
											    multipart.addBodyPart(messageBodyPart);

											    if (attachment)
											    {
											        messageBodyPart = new MimeBodyPart();
											        String filename = "test.jpg";
											        DataSource source = new FileDataSource(filename);
											        messageBodyPart.setDataHandler(new DataHandler(source));
											        messageBodyPart.setFileName(filename);
											        multipart.addBodyPart(messageBodyPart);
											    }

											    mimeReply.setContent(multipart);

											    Transport transport = getMailSession.getTransport("smtp");
											    transport.connect(host, username, pass);
											    transport.sendMessage(mimeReply, InternetAddress.parse(replyTo));
											    transport.close();
											    System.out.println("Message Sent!");

						             */
						               


						            
						            model2.setValueAt("Unread email read and replied successfully  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 );

						          }else{
						        	  model2.setValueAt("Unread email read  successfully  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 );
						          }
						          model2.setValueAt("All unread emails have been read and replied", (id-1),4 );
						      
						          }
						          model2.setValueAt("Tasks completed for this account", (id-1),4 ); 
									}catch(Exception gh)
									{
										gh.printStackTrace();
										 model2.setValueAt("Operation failed due to incorrect login details", (id-1),4 );
										 
										 f_l_l=1;
										 f_l = f_l+f_l_l;
									}
						          
									}
									
									if(table_user.contains("aol.com"))
									{	
										
										try{
										
						          store.connect("imap.aol.com",table_user, table_pass);
						        
						          model2.setValueAt("logged in successfully", (id-1),4 );
									System.out.println("logged in successfully :" +(id-1));

									  low = 1000;
									  high = 5000;
									 
									  ran   = r.nextInt(high-low) + low;
									 Thread.sleep(ran);
                                      s_l_l=1;
							         
							         s_l = s_l+s_l_l;
						          // store.connect("imap.aol.com","shemikamoloney653237@gmail.com", "lkpxkklueo");
						          folder = (IMAPFolder) store.getFolder("Inbox"); // This doesn't work for other email account
						          folder_spam = (IMAPFolder) store.getFolder("Bulk Mail");
						          Folder[] f = store.getDefaultFolder().list();
						        		  
						          for(Folder fd:f)
						              System.out.println(">> "+fd.getName());
						          
						          //folder = (IMAPFolder) store.getFolder("inbox"); This works for both email account
						          model2.setValueAt("Unmarking spam messages as not spam", (id-1),4 );
						          if(!folder_spam.isOpen())
							          folder_spam.open(Folder.READ_WRITE);
							          Message[] messages_spam = folder_spam.getMessages();
						          
							          System.out.println("No of Spam Messages found: " + folder_spam.getMessageCount());
							          model2.setValueAt("No of Spam Messages found : " + folder_spam.getMessageCount(), (id-1),4 );
							          
							          t_s_l = folder_spam.getMessageCount();
							          t_s=t_s+t_s_l;
						          int sp = folder_spam.getMessageCount();
						          if (sp != 0) {
						        	  try{
						              folder_spam.copyMessages(messages_spam, folder);
						              folder_spam.setFlags(messages_spam, new Flags(Flags.Flag.DELETED), true);
						        	  }catch(Exception tr)
						        	  {
						        		  
						        	  }
						              // Dump out the Flags of the moved messages, to insure that
						              // all got deleted

									  low = 1000;
									  high = 5000;
									 
									  ran   = r.nextInt(high-low) + low;
									 Thread.sleep(ran);
						        	  try{
						              for (int i = 0; i < sp; i++) {
						                if (!messages_spam[i].isSet(Flags.Flag.DELETED))
						                  System.out.println("Message # " + messages_spam[i] + " not deleted");
						              }
						        	  }catch(Exception rt)
						        	  {
						        		  
						        	  }
						        	  
						        	  model2.setValueAt("All spam messages have been marked as not spam!", (id-1),4 );
						            }
						          
						         

						          if(!folder.isOpen())
						          folder.open(Folder.READ_WRITE);
						          Message[] messages = folder.getMessages();
						        //  System.out.println("No of Messages : " + folder.getMessageCount());
						          System.out.println("No of Unread Messages : " + folder.getUnreadMessageCount());
						          model2.setValueAt("No of Unread Messages : " + folder.getUnreadMessageCount(), (id-1),4 );
						          
						          t_r_l = folder.getUnreadMessageCount();
						          t_r = t_r +  t_r_l ;
						        
						          
						          System.out.println(messages.length);
						          
						       // Fetch unseen messages from inbox folder
						          
						          model2.setValueAt("Getting unread emails from inbox", (id-1),4 );
						          
						          
						          Message[] messages2 = folder.search(
						              new FlagTerm(new Flags(Flags.Flag.SEEN), false));
						          /*
						          try{
						          Arrays.sort( messages2, ( m1, m2 ) -> {
						              try {
						                return m2.getSentDate().compareTo( m1.getSentDate() );
						              } catch ( MessagingException e ) {
						                throw new RuntimeException( e );
						              }
						            } );
						          }catch(Exception fg)
						          {
						        	  
						          }
						          */
						          model2.setValueAt("Reading  all unread emails from inbox", (id-1),4 );
						          int un = folder.getUnreadMessageCount();
						          for (int i=0; i < un;i++) 
						          {
						        	  model2.setValueAt("Reading   unread email  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 );

						            System.out.println("*****************************************************************************");
						            System.out.println("MESSAGE " + (i + 1) + ":");
						            //reads only unseen messages
						            Message msg =  messages2[i];
						            
						            
						            //reads all the messages in inbox
						       //    Message msg =  messages[i];
						            
						            //System.out.println(msg.getMessageNumber());
						            //Object String;
						            //System.out.println(folder.getUID(msg)

						            subject = msg.getSubject();
						            
						          

						            System.out.println("Subject: " + subject);
						            System.out.println("From: " + msg.getFrom()[0]);
						           System.out.println("To: "+msg.getAllRecipients()[0]);
						            System.out.println("Date: "+msg.getReceivedDate());
						            System.out.println("Size: "+msg.getSize());
						            System.out.println(msg.getFlags());
						            System.out.println("Body: \n"+ msg.getContent());
						            System.out.println(msg.getContentType());
						            
						          //  msg.setFlag(Flags.Flag.DELETED, true);
						            
						        //    msg.getFolder().copyMessages(messages2, folder_spam);
						            
						        //    folder.copyMessages(messages, folder_spam); //Moves email msg to SPAM folder
						            
						            msg.setFlag(Flags.Flag.SEEN, true);
						            
						            if(Reply == true)
							        	  
							          {
						           
						            Date     date = msg.getSentDate();
						               // Get all the information from the message
						               String from = InternetAddress.toString(msg.getFrom());
						               
						               
						            
						               if (from != null) {
						                  System.out.println("From: " + from);
						               }
						               String replyTo = InternetAddress.toString(msg
							         .getReplyTo());
						               if (replyTo != null) {
						                  System.out.println("Reply-to: " + replyTo);
						               }
						               String to = InternetAddress.toString(msg
							         .getRecipients(Message.RecipientType.TO));
						               if (to != null) {
						                  System.out.println("To: " + to);
						               }

						               String subject2 = msg.getSubject();
						               if (subject2 != null) {
						                  System.out.println("Subject: " + subject2);
						               }
						               Date sent = msg.getSentDate();
						               if (sent != null) {
						                  System.out.println("Sent: " + sent);
						               }
						               
						               
						               
						               InputStream inp = new FileInputStream(path);
										 Workbook wb = WorkbookFactory.create(inp);
										 Sheet sheet = wb.getSheetAt(0);
										 int ctr  = 1;
										 Row row = null;
									      Cell cell =null; 
									      
									      row = sheet.getRow(id);
									      
									      cell = row.getCell(4);
									      
									      if(cell == null)
									      {
									    	  cell = row.createCell(4);
									      }
										
						               
						               String from2=table_user;//change accordingly
						               String password=table_pass;//change accordingly

						               System.out.println("\n 1st ===> setup Mail Server Properties..");
										mailServerProperties = System.getProperties();
										mailServerProperties.put("mail.smtp.host", "smtp.aol.com");
										mailServerProperties.put("mail.smtp.socketFactory.port", "465");
										mailServerProperties.put("mail.smtp.socketFactory.class",
							               "javax.net.ssl.SSLSocketFactory");
										mailServerProperties.put("mail.smtp.auth", "true");
										mailServerProperties.put("mail.smtp.port", " 587");
										mailServerProperties.put("mail.debug", "true");
										mailServerProperties.put("mail.store.protocol", "imap");
										mailServerProperties.put("mail.transport.protocol", "smtp");
							               
										mailServerProperties.put("mail.smtp.starttls.enable", "true");
									
							               
									//	mailServerProperties.put("mail.smtp.starttls.enable", "true");
										
								     	try{
											String proxy_u =table_proxy2;
											String[] p = proxy_u.split (":");
										 ip=p[0].trim();
										
										 port=p[1].trim();
										
										 mailServerProperties.setProperty("socksProxyHost",ip); 
								            
										 mailServerProperties.setProperty("socksProxyPort",port);
										
											}catch(Exception ds)
											{
												
											} 
								     	getMailSession = Session.getDefaultInstance(mailServerProperties, null);
						               
						               
						               //Get the session object
						               Properties props = new Properties();
						               props.put("mail.smtp.host", "smtp.aol.com");
						               props.put("mail.smtp.socketFactory.port", "465");
						               props.put("mail.smtp.socketFactory.class",
						               "javax.net.ssl.SSLSocketFactory");
						               props.put("mail.smtp.auth", "true");
						               props.put("mail.smtp.port", "465");
						               props.put("mail.debug", "true");
						               props.put("mail.store.protocol", "imap");
						               props.put("mail.transport.protocol", "smtp");
						               
						               props.put("mail.smtp.starttls.enable", "true");
						               try{
											String proxy_u =table_proxy2;
											String[] p = proxy_u.split (":");
										 ip=p[0].trim();
										
										 port=p[1].trim();
										
										 props.setProperty("socksProxyHost",ip); 
								            
										 props.setProperty("socksProxyPort",port);
										
											}catch(Exception ds)
											{
												
											} 
						               /*
						               Session session2 = Session.getInstance(props,
						               new javax.mail.Authenticator() {
						               protected PasswordAuthentication getPasswordAuthentication() {
						               return new PasswordAuthentication(from2,password);
						               }
						               });

						               //compose message
						          
						               MimeMessage reply = (MimeMessage)msg.reply(false);
										reply.setFrom(new InternetAddress(from2));
										reply.setText(cell.toString());
									

										Transport.send(reply,from2,password);
						               System.out.println("message replied successfully ....");
						               */
						               MimeMessage reply = (MimeMessage)msg.reply(false);
										reply.setFrom(new InternetAddress(from2));
										reply.setText(cell.toString());
										  Session session2 = Session.getInstance(mailServerProperties,
									               new javax.mail.Authenticator() {
									               protected PasswordAuthentication getPasswordAuthentication() {
									               return new PasswordAuthentication(from2,password);
									               }
									               });

										Transport.send(reply,from2,password);
						                  System.out.println("message replied successfully ....");
						                  ++t_u_r_l;
						                  t_u_r=t_u_r+t_u_r_l;


						            
						            model2.setValueAt("Unread email read and replied successfully  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 );

						          }else{
						        	  model2.setValueAt("Unread email read  successfully  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 );  
						          }
						          model2.setValueAt("All unread emails have been read and replied", (id-1),4 );
						      
						          }
						          model2.setValueAt("Tasks completed for this account", (id-1),4 );
										}catch(Exception fh)
										{
											fh.printStackTrace();
											 model2.setValueAt("Operation failed due to incorrect login details", (id-1),4 );
											 
											 f_l_l=1;
											 f_l = f_l+f_l_l;
										}
									}
									
									if(table_user.contains("hotmail.com"))
									{	
										
										try{
										
						          store.connect("imap-mail.outlook.com",table_user, table_pass);
						        
						          model2.setValueAt("logged in successfully", (id-1),4 );
									System.out.println("logged in successfully :" +(id-1));
									
                                      s_l_l=1;
							         
							         s_l = s_l+s_l_l;

									  low = 1000;
									  high = 5000;
									 
									  ran   = r.nextInt(high-low) + low;
									 Thread.sleep(ran);
						          // store.connect("imap.aol.com","shemikamoloney653237@gmail.com", "lkpxkklueo");
						          folder = (IMAPFolder) store.getFolder("Inbox"); // This doesn't work for other email account
						          folder_spam = (IMAPFolder) store.getFolder("Junk");
						          Folder[] f = store.getDefaultFolder().list();
						        		  
						          for(Folder fd:f)
						              System.out.println(">> "+fd.getName());
						          
						          //folder = (IMAPFolder) store.getFolder("inbox"); This works for both email account
						          model2.setValueAt("Unmarking spam messages as not spam", (id-1),4 );
						          if(!folder_spam.isOpen())
							          folder_spam.open(Folder.READ_WRITE);
							          Message[] messages_spam = folder_spam.getMessages();
						          
							          System.out.println("No of Spam Messages found: " + folder_spam.getMessageCount());
							          model2.setValueAt("No of Spam Messages found : " + folder_spam.getMessageCount(), (id-1),4 );
							          
							          t_s_l = folder_spam.getMessageCount();
							          t_s=t_s+t_s_l;
						          int sp = folder_spam.getMessageCount();
						          if (sp != 0) {
						        	  try{
						              folder_spam.copyMessages(messages_spam, folder);
						              folder_spam.setFlags(messages_spam, new Flags(Flags.Flag.DELETED), true);
						        	  }catch(Exception tr)
						        	  {
						        		  
						        	  }
						              // Dump out the Flags of the moved messages, to insure that
						              // all got deleted
						        	  

									  low = 1000;
									  high = 5000;
									 
									  ran   = r.nextInt(high-low) + low;
									 Thread.sleep(ran);
						        	  try{
						              for (int i = 0; i < sp; i++) {
						                if (!messages_spam[i].isSet(Flags.Flag.DELETED))
						                  System.out.println("Message # " + messages_spam[i] + " not deleted");
						              }
						        	  }catch(Exception rt)
						        	  {
						        		  
						        	  }
						        	  
						        	  model2.setValueAt("All spam messages have been marked as not spam!", (id-1),4 );
						            }

						         
						          if(!folder.isOpen())
						          folder.open(Folder.READ_WRITE);
						          Message[] messages = folder.getMessages();
						        //  System.out.println("No of Messages : " + folder.getMessageCount());
						          System.out.println("No of Unread Messages : " + folder.getUnreadMessageCount());
						          model2.setValueAt("No of Unread Messages : " + folder.getUnreadMessageCount(), (id-1),4 );
						          
						          t_r_l = folder.getUnreadMessageCount();
						          t_r = t_r +  t_r_l ;
						        
						          
						          System.out.println(messages.length);
						          
						       // Fetch unseen messages from inbox folder
						          
						          model2.setValueAt("Getting unread emails from inbox", (id-1),4 );
						          
						          Message[] messages2 = folder.search(
						              new FlagTerm(new Flags(Flags.Flag.SEEN), false));
						          /*
						          Arrays.sort( messages2, ( m1, m2 ) -> {
						              try {
						                return m2.getSentDate().compareTo( m1.getSentDate() );
						              } catch ( MessagingException e ) {
						                throw new RuntimeException( e );
						              }
						            } );
						            */
						          model2.setValueAt("Reading  all unread emails from inbox", (id-1),4 );
						          int un = folder.getUnreadMessageCount();
						          for (int i=0; i < un;i++) 
						          {
						        	  model2.setValueAt("Reading   unread email  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 );

						            System.out.println("*****************************************************************************");
						            System.out.println("MESSAGE " + (i + 1) + ":");
						            //reads only unseen messages
						            Message msg =  messages2[i];
						            
						            
						            //reads all the messages in inbox
						       //    Message msg =  messages[i];
						            
						            //System.out.println(msg.getMessageNumber());
						            //Object String;
						            //System.out.println(folder.getUID(msg)

						            subject = msg.getSubject();
						            
						          

						            System.out.println("Subject: " + subject);
						            System.out.println("From: " + msg.getFrom()[0]);
						           System.out.println("To: "+msg.getAllRecipients()[0]);
						            System.out.println("Date: "+msg.getReceivedDate());
						            System.out.println("Size: "+msg.getSize());
						            System.out.println(msg.getFlags());
						            System.out.println("Body: \n"+ msg.getContent());
						            System.out.println(msg.getContentType());
						            
						          //  msg.setFlag(Flags.Flag.DELETED, true);
						            
						        //    msg.getFolder().copyMessages(messages2, folder_spam);
						            
						        //    folder.copyMessages(messages, folder_spam); //Moves email msg to SPAM folder
						            
						            msg.setFlag(Flags.Flag.SEEN, true);
						            
						            if(Reply == true)
							          {
						           
						            Date     date = msg.getSentDate();
						               // Get all the information from the message
						               String from = InternetAddress.toString(msg.getFrom());
						               
						               
						            
						               if (from != null) {
						                  System.out.println("From: " + from);
						               }
						               String replyTo = InternetAddress.toString(msg
							         .getReplyTo());
						               if (replyTo != null) {
						                  System.out.println("Reply-to: " + replyTo);
						               }
						               String to = InternetAddress.toString(msg
							         .getRecipients(Message.RecipientType.TO));
						               if (to != null) {
						                  System.out.println("To: " + to);
						               }

						               String subject2 = msg.getSubject();
						               if (subject2 != null) {
						                  System.out.println("Subject: " + subject2);
						               }
						               Date sent = msg.getSentDate();
						               if (sent != null) {
						                  System.out.println("Sent: " + sent);
						               }
						               
						               
						               
						               InputStream inp = new FileInputStream(path);
										 Workbook wb = WorkbookFactory.create(inp);
										 Sheet sheet = wb.getSheetAt(0);
										 int ctr  = 1;
										 Row row = null;
									      Cell cell =null; 
									      
									      row = sheet.getRow(id);
									      
									      cell = row.getCell(4);
									      
									      if(cell == null)
									      {
									    	  cell = row.createCell(4);
									      }
										
						               
									      System.out.println("\n 1st ===> setup Mail Server Properties..");
											mailServerProperties = System.getProperties();
											mailServerProperties.put("mail.smtp.host", "smtp-mail.outlook.com");
											mailServerProperties.put("mail.smtp.socketFactory.port", "465");
											mailServerProperties.put("mail.smtp.socketFactory.class",
								               "javax.net.ssl.SSLSocketFactory");
											mailServerProperties.put("mail.smtp.auth", "true");
											mailServerProperties.put("mail.smtp.port", "587");
											mailServerProperties.put("mail.debug", "true");
											mailServerProperties.put("mail.store.protocol", "imap");
											mailServerProperties.put("mail.transport.protocol", "smtp");
								               
											mailServerProperties.put("mail.smtp.starttls.enable", "true");
										
								               
										//	mailServerProperties.put("mail.smtp.starttls.enable", "true");
											
									     	try{
												String proxy_u =table_proxy2;
												String[] p = proxy_u.split (":");
											 ip=p[0].trim();
											
											 port=p[1].trim();
											
											 mailServerProperties.setProperty("socksProxyHost",ip); 
									            
											 mailServerProperties.setProperty("socksProxyPort",port);
											
												}catch(Exception ds)
												{
													
												} 
									     	getMailSession = Session.getDefaultInstance(mailServerProperties, null);
						               String from2=table_user;//change accordingly
						               String password=table_pass;//change accordingly

						               //Get the session object
						               Properties props = new Properties();
						               props.put("mail.smtp.host", "smtp-mail.outlook.com");
						               props.put("mail.smtp.socketFactory.port", "465");
						               props.put("mail.smtp.socketFactory.class",
						               "javax.net.ssl.SSLSocketFactory");
						               props.put("mail.smtp.auth", "true");
						               props.put("mail.smtp.port", "587");
						               props.put("mail.debug", "true");
						               props.put("mail.store.protocol", "imap");
						               props.put("mail.transport.protocol", "smtp");
						               
						               props.put("mail.smtp.starttls.enable", "true");
						               try{
											String proxy_u =table_proxy2;
											String[] p = proxy_u.split (":");
										 ip=p[0].trim();
										
										 port=p[1].trim();
										
										 props.setProperty("socksProxyHost",ip); 
								            
										 props.setProperty("socksProxyPort",port);
										
											}catch(Exception ds)
											{
												
											} 
						               /*
						               Session session2 = Session.getInstance(props,
						               new javax.mail.Authenticator() {
						               protected PasswordAuthentication getPasswordAuthentication() {
						               return new PasswordAuthentication(from2,password);
						               }
						               });*/

						               //compose message
						           	Message replyMessage = new MimeMessage(getMailSession);
					                  replyMessage = (MimeMessage) msg.reply(false);
					                  replyMessage.setFrom(new InternetAddress(to));
					                  replyMessage.setText(cell.toString());
					                  replyMessage.setReplyTo(msg.getReplyTo());

					                  // Send the message by authenticating the SMTP server
					                  // Create a Transport instance and call the sendMessage
					                  Transport t = getMailSession.getTransport("smtp");
					                  try {
						   	     //connect to the smpt server using transport instance
							     //change the user and password accordingly	
						             t.connect("smtp-mail.outlook.com",table_user, table_pass);
						             t.sendMessage(replyMessage,
					                        replyMessage.getAllRecipients());
					                  } finally {
					                     t.close();
					                  }
					                  System.out.println("message replied successfully ....");
					                  ++t_u_r_l;
					                  t_u_r=t_u_r+t_u_r_l;


						            
						            model2.setValueAt("Unread email read and replied successfully  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 );

						          }else{
						        	  model2.setValueAt("Unread email read successfully  of "+(i + 1)+" of "+un+" from inbox", (id-1),4 ); 
						          }
						          model2.setValueAt("All unread emails have been read and replied", (id-1),4 );
						      
						          }
						          model2.setValueAt("Tasks completed for this account", (id-1),4 );
										}catch(Exception fh)
										{
											fh.printStackTrace();
											 model2.setValueAt("Operation failed possibly due to incorrect login details", (id-1),4 );
											 
											 f_l_l=1;
											 f_l = f_l+f_l_l;
										}
									}
									
									
						        }
						        finally 
						        {
						          if (folder != null && folder.isOpen()) { folder.close(true); }
						          if (store != null) { store.close(); }
						        }

							  }catch(Exception gtr)
							  {
								 gtr.printStackTrace();
							  }
						 
						
						 /*
								}
								else{
									y=0;
									
									try{
										model2.setValueAt("waiting for scheduled time", (id-1),4 );
										
										System.out.println("Bot is waiting to be started by the given hour and minutes");
										/*
										st="Bot is waiting to be started by the given hour and minutes";
										
										update(st);
										
									Thread.sleep(1000);*/
						 
						 /*
									}catch(Exception ba)
									
									{
										
									}
								}
							*/
						 break;
				
								
						 }
							return null;
					}
			
				};
				email.execute();
				
	}
	
	public static void bet2(){
		SwingWorker<Void,Void> email = new SwingWorker<Void,Void>()
				{

					@Override
					protected Void doInBackground() throws Exception {
						
						
				
						 try{
							 
							 for(int i=0;i<10;i++)
							 {
								 
								 textField.setText(String.valueOf(s_l));
								 
								 textField_3.setText(String.valueOf(f_l));
									 
								 textField_4.setText(String.valueOf(t_s));
								 
								 textField_5.setText(String.valueOf(t_r));
								 
								 textField_6.setText(String.valueOf(t_u_r));
								 
									
								 Thread.sleep(1000);
								 i=0;
							 }
						         

							  }catch(Exception gtr)
							  {
								 gtr.printStackTrace();
							  }
						
							return null;
					}
			
				};
				email.execute();
				
	}
	
	public static void bet5(){
		SwingWorker<Void,Void> email = new SwingWorker<Void,Void>()
				{

					@Override
					protected Void doInBackground() throws Exception {
						
						
				
						 try{
							 
							 for(int i=0;i<10;i++)
							 {
								 
									DayOfWeek dayOfWeek2 = DayOfWeek.from(LocalDate.now());
									
					        		
					        		String d = String.valueOf(dayOfWeek2);
					        	//	System.out.println(d);
					        		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH:mm");  
						        	   LocalDateTime now = LocalDateTime.now(); 
						        	   String da=dtf.format(now);
						        	 //  System.out.println(da);  
						        	   
						        	  // textField_5.setText(d+" "+da);
								
						        	   System.out.println(d+" "+da);
						        	   
						        	   String [] ho = da.split(":");
						        	   String h =ho[0];
						        	   String m =ho[1];
						        	   nike_2();
						        	   
						        	   Thread.sleep(600000);
						        	   
						        	   /*
									
									if((hour_j <= Integer.parseInt(h)) &&(min_j <= Integer.parseInt(m)))
										
									{
										label.setText("Time's up, accessing mail accounts");
										
										 s_l=0;
										 f_l=0;
										 t_s=0;
									     t_r=0;
									     t_u_r=0;
									     
									//	nike_2();
										
										Thread.sleep(600000);
									}else{
										label.setText("Waiting for the scheduled time");
									}
									*/
								 
									
								 Thread.sleep(1000);
								 i=0;
							 }
						         

							  }catch(Exception gtr)
							  {
								 gtr.printStackTrace();
							  }
						
							return null;
					}
			
				};
				email.execute();
				
	}
	
	public static void bet6(){
		SwingWorker<Void,Void> email = new SwingWorker<Void,Void>()
				{

					@Override
					protected Void doInBackground() throws Exception {
						
						
						try {

							for(int i=0;i<10;i++)
							{
								
							//	sr();
								
							 String value1 = comboBox.getSelectedItem().toString();
								Thread.sleep(Integer.parseInt(value1)*3600000);
							// Thread.sleep(Integer.parseInt(value1)*60000);
								nike_2();
								
								i=0;
							}
									
									 
						
						} catch (Exception y) {
							y.printStackTrace();

							StringWriter sw = new StringWriter();
							y.printStackTrace(new PrintWriter(sw));
							String exceptionAsString = sw.toString();
							System.out.println(exceptionAsString);
							JOptionPane.showMessageDialog(null, "" + exceptionAsString + "", "ERROR!",
									JOptionPane.ERROR_MESSAGE);

							// driver2.close();
							// driver2.quit();
							// Thread.sleep(10000);
							// rebeet();
						}
						
							return null;
					}
			
				};
				email.execute();
				
	}
	
	public static void bet7(){
		SwingWorker<Void,Void> email = new SwingWorker<Void,Void>()
				{

					@Override
					protected Void doInBackground() throws Exception {
						
						
						try {

							for(int i=0;i<10;i++)
							{
								
							//	sr();
								
							 String value1 = comboBox.getSelectedItem().toString();
								Thread.sleep(Integer.parseInt(value1)*3600000);
								nike_3();
								
								i=0;
							}
									
									 
						
						} catch (Exception y) {
							y.printStackTrace();

							StringWriter sw = new StringWriter();
							y.printStackTrace(new PrintWriter(sw));
							String exceptionAsString = sw.toString();
							System.out.println(exceptionAsString);
							JOptionPane.showMessageDialog(null, "" + exceptionAsString + "", "ERROR!",
									JOptionPane.ERROR_MESSAGE);

							// driver2.close();
							// driver2.quit();
							// Thread.sleep(10000);
							// rebeet();
						}
						
							return null;
					}
			
				};
				email.execute();
				
	}
	
	public static void nike_2(){
		SwingWorker<Void,Void> email = new SwingWorker<Void,Void>()
				{

					@Override
					protected Void doInBackground() throws Exception {
					
						try{
							InputStream inp = new FileInputStream(path);
							 Workbook wb = WorkbookFactory.create(inp);
							 Sheet sheet = wb.getSheetAt(0);
							 int ctr  = 1;
							 Row row = null;
						      Cell cell =null; 
						      
						      
						  	String f =	textField_1.getText();
							
							int f_i = Integer.parseInt(f);  
							
							String f2 =	textField_2.getText();
							
							int f2_i = Integer.parseInt(f2);
							
							f2_i=f2_i+1;
						      for(f_i=f_i;f_i<f2_i;f_i++){
						    	  row = sheet.getRow(f_i);
						    	//    cell = row.getCell(0);
						    	    
						    	    double id = row.getCell(0).getNumericCellValue();
						    	    int id_int =(int)id;
						    	    Cell user = row.getCell(1);
						    	    if(user == null)
						    	    {
						    	    	user = row.createCell(1);
						    	    }
						    	    Cell pass = row.getCell(2);
						    	    
						    	    if(pass == null)
						    	    {
						    	    	pass = row.createCell(2);
						    	    }
				                //   double   excelSubject = row.getCell(3).getNumericCellValue();
				                   Cell proxy = row.getCell(3);
				                   
				                   if(proxy == null)
						    	    {
						    	    	proxy = row.createCell(3);
						    	    }
				                   
				                   
				              /*     Cell HEAD = row.getCell(5);
				                    
				                    
				                   int cv =(int)excelSubject;
				                   
				                   String cv_s = String.valueOf(cv);
				                   */
				               //    nike_checkout(excelName.toString(),excelGender.toString(),excelProgrammingLanguage.toString(),cv_s,excelSubject2.toString(),i,HEAD.toString());
				                   
				                   bet((id_int) ,user.toString() ,pass.toString(),proxy.toString()  );  
				                   
				                   Thread.sleep(10000);
						    	    
						      }
						      bet6();
							
						}catch(Exception h)
						{
							h.printStackTrace();
							JOptionPane.showMessageDialog(null, h.getMessage());
						}
						
						
						
							return null;
					}
			
				};
				email.execute();
				
	}
	
	public static void nike_3(){
		SwingWorker<Void,Void> email = new SwingWorker<Void,Void>()
				{

					@Override
					protected Void doInBackground() throws Exception {
					
						try{
							InputStream inp = new FileInputStream(path);
							 Workbook wb = WorkbookFactory.create(inp);
							 Sheet sheet = wb.getSheetAt(0);
							 int ctr  = 1;
							 Row row = null;
						      Cell cell =null; 
						      
						      
						
						      for(int i=1; i<= sheet.getLastRowNum(); i++){
						    	  row = sheet.getRow(i);
						    	//    cell = row.getCell(0);
						    	    
						    	    double id = row.getCell(0).getNumericCellValue();
						    	    int id_int =(int)id;
						    	    Cell user = row.getCell(1);
						    	    if(user == null)
						    	    {
						    	    	user = row.createCell(1);
						    	    }
						    	    Cell pass = row.getCell(2);
						    	    
						    	    if(pass == null)
						    	    {
						    	    	pass = row.createCell(2);
						    	    }
				                //   double   excelSubject = row.getCell(3).getNumericCellValue();
				                   Cell proxy = row.getCell(3);
				                   
				                   if(proxy == null)
						    	    {
						    	    	proxy = row.createCell(3);
						    	    }
				              /*     Cell HEAD = row.getCell(5);
				                    
				                    
				                   int cv =(int)excelSubject;
				                   
				                   String cv_s = String.valueOf(cv);
				                   */
				               //    nike_checkout(excelName.toString(),excelGender.toString(),excelProgrammingLanguage.toString(),cv_s,excelSubject2.toString(),i,HEAD.toString());
				                   
				                   bet((id_int) ,user.toString() ,pass.toString(),proxy.toString()  );  
				                   
				                   Thread.sleep(10000);
						    	    
						      }
						      bet7();
							
						}catch(Exception h)
						{
							h.printStackTrace();
							JOptionPane.showMessageDialog(null, h.getMessage());
						}
						
						
						
							return null;
					}
			
				};
				email.execute();
				
	}
	  public static void user(boolean t)  {
	    	 
	    	userN=t;
	    	

	  	}
}
