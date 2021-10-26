public class Record {

    private int row;
    private int year;
    private String month;
    private String[] data;

    public Record() {

    }

    public Record(int row, int year, String month, String[] data) {
        this.row = row;
        this.year = year;
        this.month = month;
        this.data = data;
    }


    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getYear() {
        return year;
    }

    public void setYear(int year) {
        this.year = year;
    }

    public String getMonth() {
        return month;
    }

    public void setMonth(String month) {
        this.month = month;
    }

    public String[] getData() {
        return data;
    }

    public void setData(String[] data) {
        this.data = data;
    }
}
