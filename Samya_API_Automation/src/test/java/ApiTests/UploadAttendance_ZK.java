package ApiTests;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.security.PublicKey;
import java.util.Properties;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.*;



import Utils.RestUtil;
import Utils.XLReader;
import io.restassured.RestAssured;
import io.restassured.http.ContentType;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;

public class UploadAttendance_ZK {
	public String zkdeviceurl = null;
	public static  String filepath= System.getProperty("user.dir")+"\\src\\testdata\\TestData.xlsx"; 
	static Sheet sheet;
	static Workbook book;
	Properties prop = null;
	
	private Response res = null;
	public  RequestSpecification request = null;
	
	FileInputStream fs = null;
			
	UploadAttendance_ZK()
	{
		prop = new Properties();
		try {
			fs = new FileInputStream(System.getProperty("user.dir")+"/src/config/config.properties");
			prop.load(fs);
			zkdeviceurl = prop.getProperty("urlfordevice");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
  
  @BeforeTest
  public void beforeTest() {
	 
	   RestUtil.setBaseURI(zkdeviceurl);
	  //RestAssured.baseURI="http://staging-bmd.securtime.in";
	  RestUtil.setContentType(ContentType.TEXT);
	   
	  
  }
  
  public static Object[][] getTestData(String sheetName){
	  
		InputStream inputStream = null;
		 try {
			 
			 inputStream = new FileInputStream(filepath);
			// System.out.println(filepath);
			  book =  WorkbookFactory.create(inputStream);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 
			
		
		 sheet = book.getSheet(sheetName);
		 
		 Object[][] data = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		      for (int i = 0; i < sheet.getLastRowNum();i++) {
		    	  for (int j = 0; j < sheet.getRow(0).getLastCellNum(); j++) {
		    		  data[i][j]=sheet.getRow(i+1).getCell(j).toString();
					
				}
				
			}
		      
		 return data;
  }
  

  
  @DataProvider
	public Object[][] getData() throws IOException{
		
		
		Object data[][]=getTestData("AttendancePunches_ZK");
		
		return data;
		
		
	}
  /*
  @DataProvider(name ="getData")
	public Object[][] PunchData(){
		//TestUtil utility=new TestUtil();
		System.out.println("sumaiya");
		//Object data[][]=utility.getTestData(sheetname);
		//return data;
	  return new Object[][]{{"115","2017-12-23","3576164500070","13:40:23"},{"115","2017-12-23","3576164500070","13:50:23"}};
		
		
	}*/
	
	
  @Test(dataProvider="getData")
   public void uploadAttendance(String enrollid,String date,String serialno,String time) {
	  
	  
	  String requestbody = enrollid+"\t"+date+" "+time+"\t0\t1\t0\t0";	
	 // System.out.println(requestbody);
	 if((serialno!= null || serialno!="") &&(enrollid!= null || enrollid!="") && (date!= null || date!="") && (time!= null || time!=""))
	 {
	  Response res = RestAssured.given().
		      queryParam("SN", serialno).
		      queryParam("table","ATTLOG").
		      queryParam("Stamp","2014-12-12").
		       body(requestbody).
		      when().
		      post("/iclock/cdata");
	  
	  
	  
	  int statuscode = res.getStatusCode();
	  System.out.println(statuscode);
	 }
	  
  }
  

@AfterTest
  public void AfterTest() {
	  
	  RestUtil.resetBaseURI();
  }
  
  

}
