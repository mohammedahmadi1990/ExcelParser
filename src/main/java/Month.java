public class Month {

    private String[] english;
    private String[] albanian;
    private String[] french;
    private int[] number;

    public Month() {

        number = new int[]{1,2,3,4,5,6,7,8,9,10,11,12};
        english = new String[]{"january","february","march","april", "may", "june", "july", "august","spetember", "october", "november", "december"};
        albanian = new String[]{"janar", "shkurt", "mars", "prill", "maj", "qershor", "korrik", "gusht", "shtator", "tetor", "nëntor", "dhjetor"};
        french = new String[]{"janvier","février","mars","avril", "mai", "juin", "juillet", "août","septembre", "octobre", "novembre", "décembre"};
    }

    public String englishToAlbanian(String text){
        for (int i = 0; i < english.length; i++) {
            if(text.toLowerCase().equals(english[i])){
                return albanian[i];
            }
        }
        return null;
    }

    public String albanianToEnglish(String text){
        for (int i = 0; i < albanian.length; i++) {
            if(text.toLowerCase().equals(albanian[i])){
                return english[i];
            }
        }
        return null;
    }

    public int AlbanianMonthToNum(String text){
        for (int i = 0; i < albanian.length; i++) {
            if(text.toLowerCase().equals(albanian[i])){
                return number[i];
            }
        }
        return 0;
    }

    public String[] getEnglish() {
        return english;
    }

    public void setEnglish(String[] english) {
        this.english = english;
    }

    public String[] getAlbanian() {
        return albanian;
    }

    public void setAlbanian(String[] albanian) {
        this.albanian = albanian;
    }

    public String[] getFrench() {
        return french;
    }

    public void setFrench(String[] french) {
        this.french = french;
    }

    public int[] getNumber() {
        return number;
    }

    public void setNumber(int[] number) {
        this.number = number;
    }
}
