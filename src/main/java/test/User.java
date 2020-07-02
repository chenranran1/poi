package test;

public class User {

	private String name;
	private String gender;
	private String age;
	private String idcard;
	private String address;
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getGender() {
		return gender;
	}
	public void setGender(String gender) {
		this.gender = gender;
	}
	public String getAge() {
		return age;
	}
	public void setAge(String age) {
		this.age = age;
	}
	public String getIdcard() {
		return idcard;
	}
	public void setIdcard(String idcard) {
		this.idcard = idcard;
	}
	public String getAddress() {
		return address;
	}
	public void setAddress(String address) {
		this.address = address;
	}
	public User(String name, String gender, String age, String idcard, String address) {
		super();
		this.name = name;
		this.gender = gender;
		this.age = age;
		this.idcard = idcard;
		this.address = address;
	}
	public User() {
		super();
	}
}
