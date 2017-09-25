import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args){
        ExcelnputXlsx excelnputXlsx = new ExcelnputXlsx();
        excelnputXlsx.getListMap();
        List<Map<Object,Object>> list =  ExcelnputXlsx.getList1();
        System.out.println("-----------------------------------------------------------");
        for (int i = 0; i < list.size(); i++) {
            System.out.print(list.get(i).get("MERCHANTNAME"));
            System.out.print(list.get(i).get("MERCHANTINDUSTRY"));
            System.out.print(list.get(i).get("PersonInCharge"));
            System.out.print(list.get(i).get("contact"));
            System.out.print(list.get(i).get("email"));
            System.out.print(list.get(i).get("qrCode"));
            System.out.print(list.get(i).get("location"));
            System.out.print(list.get(i).get("salesPerson"));
            System.out.print(list.get(i).get("ID"));
            System.out.println();
            System.out.println("-------------------------");
        }

    }
}
