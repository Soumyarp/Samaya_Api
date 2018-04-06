package Utils;



import io.restassured.RestAssured;
import static io.restassured.RestAssured.*;
import io.restassured.http.ContentType;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;

public class RestUtil {
	
	//Global Setup Variables
    public static String path; //Rest request path
    public  RequestSpecification request = null;
	/***Sets Base URI***
    Before starting the test, we should set the RestAssured.baseURI
    */
    public static void setBaseURI (String baseURI){
        RestAssured.baseURI = baseURI;
    	System.out.println(baseURI);
    }
    /*
     ***Sets ContentType***
     We should set content type as JSON or XML before starting the test
     */
     public static void setContentType (ContentType Type){
         given().contentType(Type);
         
     }
    
     
     public static void resetBaseURI (){
         RestAssured.baseURI = null;
     }
}
