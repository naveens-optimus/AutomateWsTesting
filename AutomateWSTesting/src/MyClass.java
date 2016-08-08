import org.apache.xmlbeans.XmlException;

import com.eviware.soapui.support.XmlHolder;

public class MyClass{
		 public int Sum(int n1, int n2){
			 System.out.println("Parent method was called");
			 return n1+n2;
		 }
		 
		 public int num=10;
		 
//		 public int Sum(long n1, int n2){
//			 System.out.println("That method was called");
//			 return (int) (n1+n2);
//		 }
//		 
//		 public int Sum(int n1, long n2){
//			 System.out.println("3 method was called");
//			 return (int) (n1+n2);
//		 }
//		 
////		 public long Sum(int n1, int n2){
////			 return n1+n2;
////		 }
//		 
//		 public int Sum(int n1, int n2, int n3){
//			 return n1+n2;
//		 }
		 
		 
	 }