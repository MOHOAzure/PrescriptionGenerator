package medic;

public class Prescription {
  private String houseName;
  private String residentName;
  private String departmentName;
  private int progress;
  //private String doctorName;

  public Prescription() {
    super();
  }

  public String toString() {
    return "Employee [" +
      "houseName=" + houseName +
      ", residentName=" + residentName +
      ", departmentName=" + departmentName +
      ", progress=" + progress +
      //", doctorName=" + doctorName +
      "]";
  }
  public Prescription(String houseName, String residentName, String departmentName, int progress) {
    super();
    this.setHouseName(houseName);
    this.setResidentName(residentName);
    this.setDepartmentName(departmentName);
    this.setProgress(progress);
  }

  public int getProgress() {
    return progress;
  }

  public void setProgress(int progress) {
    this.progress = progress;
  }

  public String getHouseName() {
    return houseName;
  }

  public void setHouseName(String houseName) {
    this.houseName = houseName;
  }

  public String getResidentName() {
    return residentName;
  }

  public void setResidentName(String residentName) {
    this.residentName = residentName;
  }

  public String getDepartmentName() {
    return departmentName;
  }

  public void setDepartmentName(String departmentName) {
    this.departmentName = departmentName;
  }
}
