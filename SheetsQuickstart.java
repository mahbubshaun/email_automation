import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.google.common.collect.Lists;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
public class SheetsQuickstart {
   public static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
    public static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    public static final String TOKENS_DIRECTORY_PATH = "tokens";
    public static String time;
    public static String value;
    public static String value2;
    public static String value3;
    public static String value4;
    public static String value5;
    public static String value6;
    public static String value7;
    public static String value8;
    public static String value9;
    public static String value10;
    public static String value11;
    public static String value12;
    public static String value13;
    public static String userN;
    public static boolean access;
    public static String con;
    public static boolean check2;
    public static boolean check;
    public static String p_n;
    public static boolean check_last;
    public static ArrayList<String> ar2 = new ArrayList<String>();
    public static String city;
    public static String state;
    static mail ev = new mail();
    /**
     * Global instance of the scopes required by this quickstart.
     * If modifying these scopes, delete your previously saved tokens/ folder.
     */
 //   private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS_READONLY);
    public static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS_READONLY);
    public static final String CREDENTIALS_FILE_PATH = "/client2_id.json";
	public static final Object This = null;
	public static final Object // Cell values ...
                		Thiss = null;

    /**
     * Creates an authorized Credential object.
     * @param HTTP_TRANSPORT The network HTTP Transport.
     * @return An authorized Credential object.
     * @throws IOException If the credentials.json file cannot be found.
     */
    public static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
      InputStream in = SheetsQuickstart.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
    	
  //	InputStream in =   	SheetsQuickstart.class.getClassLoader().getResourceAsStream(CREDENTIALS_FILE_PATH);
    	
    	
    	
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    /*
     * Prints the names and majors of students in a sample spreadsheet:
     * https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
     */
    public static void main(String... args) throws IOException, GeneralSecurityException {
        // Build a new authorized API client service.
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        final String spreadsheetId = "1svMRxiC1X6J9gWHxp02seDH0JXCc6mfh46QieLKAG24";
        final String range = "Sheet1!A1:P";
        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();
        
        check = true;
        access = false;
        //SORTING START
        /*
        BatchUpdateSpreadsheetRequest busReq = new BatchUpdateSpreadsheetRequest();
        SortSpec ss = new SortSpec();
        // ordering ASCENDING or DESCENDING
        ss.setSortOrder("DESCENDING");
        // the column number starting from 0
        ss.setDimensionIndex(0);
        SortRangeRequest srr = new SortRangeRequest();
        srr.setSortSpecs(Arrays.asList(ss));
        Request req = new Request();
        req.setSortRange(srr);
        busReq.setRequests(Arrays.asList(req));
        // mService is a instance of com.google.api.services.sheets.v4.Sheets
        service.spreadsheets().batchUpdate(spreadsheetId, busReq).execute();
        */
        
        //SORTING END
      
		
        
        if(check2 == true)
        {
        	//write operation start
            //  @SuppressWarnings("unchecked")
              
              
              
              ValueRange oRange = new ValueRange();
              List<List<Object>> arrData = getData();
              oRange.setValues(arrData);

              AppendValuesResponse appendResult = service.spreadsheets().values().append(spreadsheetId,range,oRange ).
             		setValueInputOption("USER_ENTERED").setInsertDataOption("INSERT_ROWS").setIncludeValuesInResponse(true).execute();
             		
              
              /*
              UpdateValuesResponse result =
                      service.spreadsheets().values().update(spreadsheetId, range, oRange)
                              .setValueInputOption("USER_ENTERED")
                              .execute();
              System.out.printf("%d cells updated.", result.getUpdatedCells());	
              */
              //write operation finish
              
        	
        	
        	check2=false;
        	 System.gc(); 
        }
        
        
        
        
        if(check == true)
        {
        	
        	//read operation
            
            ValueRange response = service.spreadsheets().values()
                    .get(spreadsheetId, range)
                    .execute();
            
         int n =   service.spreadsheets().values()
            .get(spreadsheetId, range)
            .execute().getValues().size();
            
      //   System.out.println(n);
            List<List<Object>> values = response.getValues();
            
         
            if (values == null || values.isEmpty()) {
                System.out.println("No data found.");
            } else {
              //  System.out.println("link");
                
            
                for (List row : values) {
                    // Print columns A and E, which correspond to indices 0 and 4.
                   System.out.printf("%s, %s\n", row.get(0), row.get(1));
                	
                   if(row.get(0).toString().equals(userN) && row.get(1).toString().equals("true") )
                   {
                	   access = true;
                	   ev.user(access);
                	   
                	   break;
                   }
              //	 System.out.printf("%s, \n",row.get(1));
           //   	ar2.add(row.get(1).toString());
                	/*
                  if(row.get(1).equals(value6))
                  {
                	  //link exist
                	  
                	  ev.exist();
                	  System.gc(); 
                	//  break;
                  }
                  */
                	
                //    System.out.println(row.size());
                   // break;
                }
              //  ev.allarr(ar2);
            }
            
            //read operation finished
        	
        	
        	check=false;
        }
        if (check_last == true)
        {
//read operation
            
            ValueRange response = service.spreadsheets().values()
                    .get(spreadsheetId, range)
                    .execute();
            
         int n =   service.spreadsheets().values()
            .get(spreadsheetId, range)
            .execute().getValues().size();
            
      //   System.out.println(n);
         ValueRange response2 = service.spreadsheets().values()
                 .get(spreadsheetId, "Sheet1!A"+n+":P")
                 .execute();
            List<List<Object>> values = response2.getValues();
            
         
            if (values == null || values.isEmpty()) {
                System.out.println("No data found.");
            } else {
              //  System.out.println("link");
                
            
                for (List row : values) {
                    // Print columns A and E, which correspond to indices 0 and 4.
                   System.out.printf("%s, %s, %s\n", row.get(5), row.get(6) ,row.get(7));
             //      ev.exist2( row.get(5).toString(), row.get(6).toString() ,row.get(7).toString());
             	  System.gc(); 
              //	 System.out.printf("%s, \n",row.get(1));
                	
                	/*
                  if(row.get(1).equals(value6))
                  {
                	  //link exist
                	  
                	  ev.exist();
                	  System.gc(); 
                	  break;
                  }
                  */
                //    System.out.println(row.size());
                   // break;
                }
            }
            
            //read operation finished
        	
        	
            check_last=false;
        }
    }
    
    public static void assing(String t, String v , String v2,String v3,String v4,String v5 ,String country,String page,String cit,String stat)  {
 
    	time=t;
    	value=v;
    	value2=v2;
    	value3=v3;
    	value4=v4;
    	value5=v5;
    	con = country;
    	p_n = page;
    	city = cit;
    	state= stat;

  	}
    
    public static void user(String t)  {
    	 
    	userN=t;
    	

  	}
    
    public static void assing2( )  {
    	 check=true;
    	//value6=link;

  	}
    
    public static void assing3( )  {
   	 check2=true;
   	

 	}
    
    public static void assing4(String id )  {
      	 value7 = id ;
      	

    	}
    public static void assing5()  {
    	check_last =true;
     	

   	}
    public static List<List<Object>> getData ()  {

    	  List<Object> data1 = new ArrayList<Object>();
    	  data1.add (time);
    	  data1.add (value);
    	  data1.add (value2);
    	  data1.add (value3);
    	  data1.add (value4);
    	  data1.add (value5);
    	  data1.add (con);
    	  
    	  data1.add (p_n);
    	  data1.add (city);
    	  data1.add (state);
    	  List<List<Object>> data = new ArrayList<List<Object>>();
    	  data.add (data1);

    	  return data;
    	}
}