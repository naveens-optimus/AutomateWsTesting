
public class TestJava {

	public static void main(String[] args) {
		//Method over riding
		System.out.println("Sum of parent=" + new MyClass().Sum(10, 5));
		
		//System.out.println("Sum of parent=" + new MyClass().Sum(10, 5));
		
		System.out.println("Sum of child=" + new YourClass().Sum(10, 5));
		System.out.println("child=" + new YourClass().num);
		
		MyClass myClass = new YourClass();
		
		System.out.println("Sum of child class with parent reference=" + myClass.Sum(10, 5));
		System.out.println("Sum of child class with parent reference=" + myClass.num);
	}
	

}
