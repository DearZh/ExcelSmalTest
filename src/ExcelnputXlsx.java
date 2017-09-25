import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.helpers.XSSFXmlColumnPr;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.text.DateFormat;
import java.util.*;
import java.util.regex.Pattern;

public class ExcelnputXlsx {
    //商户表对应的数据字段
    private static List<Map<Object,Object>> list1 = new ArrayList<Map<Object, Object>>();
    //商户结算表所对应的数据字段
    private static List<Map<Object,Object>> list2 = new ArrayList<Map<Object, Object>>();
    //二维码表所对应的数据字段
    private static List<Map<Object,Object>> list3 = new ArrayList<Map<Object, Object>>();
    private static Map<Object,Object> mapReal = null;
    private static  Map<Object,Object> mapReal2 = null;
    private static  Map<Object,Object> mapReal3 = null;
    long ID = 0;
    long chargeID = 0;
    public List<Map<String,String>> getListMap(){
        List<Map<String,String>> list = new ArrayList<Map<String,String>>();
        Map<String,String> map = null;
        String fileNamePath = "";
        String sheetName = "";
        int countNum = 0;
        Properties properties = new Properties();
        InputStream inputStream = Main.class.getClassLoader().getResourceAsStream("ExcelSmal.properties");
        try {
            properties.load(inputStream);
            fileNamePath = properties.getProperty("filePathName");
            sheetName = properties.getProperty("sheetName");
            countNum = Integer.parseInt(properties.getProperty("countNum"));
            ID = Long.parseLong(properties.getProperty("ID"));
            chargeID = Long.parseLong(properties.getProperty("chargeID"));
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            System.out.println(fileNamePath);
            File file = new File(fileNamePath);
            InputStream inputStream1 = new FileInputStream(file);
            //读取指定位置的文件，
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(inputStream1);
            //读取指定 名称的 sheet表格 数据
            XSSFSheet xssfSheet = xssfWorkbook.getSheet(sheetName);
            //循环当前所需读取数据的所有行数
            for (int i = 1; i < countNum; i++) {
                XSSFRow xssfRow = xssfSheet.getRow(i);
                //获取当前行所对应的列数**(xssfRow获取数据从0开始)
                int cellNum =  xssfRow.getLastCellNum();
                cellNum+=1;
                mapReal = new HashMap<Object,Object>();
                mapReal2 = new HashMap<Object,Object>();
                mapReal3 = new HashMap<Object,Object>();
                for (int j = 1; j < cellNum; j++) {
                    //获取当前行所指定列的cell数据内容；
                    Object content = getJavaValue(xssfRow.getCell(j));
                    if(content!=null){
                        //表示当前数据为空
                        if(content.toString().length()==0){
                            setList(j,content);
                        }else{
                            //获取当前非空的数据
                            setList(j,content);
                        }
                    }else{
                        //获取当前非空的数据；
                        setList(j,content);
                    }
                }
                //设置商户表的自增Id+1
                ID+=1;
                chargeID+=1;
                mapReal.put("ID",ID);
                //设置商户结算表Id+1
                mapReal2.put("ID",chargeID);
                list1.add(mapReal);
                list2.add(mapReal2);
                list3.add(mapReal3);
                //System.out.println();
                //System.out.println("-------------------");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }
    //存储所获取的数据到指定的List集合中
    public void setList(int j,Object conent){
        if(j!=0){
            switch (j){
                case 1:
                    mapReal.put("MERCHANTNAME",conent);
                    break;
                case 2:
                    break;
                case 3:
                    break;
                case 4:
                    break;
                case 5:
                    break;
                case 6:
                    break;
                case 7:
                    break;
                case 8:
                    String realConent = conent.toString();
                    int startIndex = realConent.indexOf("(");
                    String databaseContent =  realConent.substring(startIndex+1,realConent.length()-1);
                    mapReal.put("MERCHANTINDUSTRY",databaseContent);
                    break;
                case 9:
                    mapReal.put("PersonInCharge",conent);
                    break;
                case 10:
                    mapReal.put("contact",conent);
                    break;
                case 11:
                    mapReal.put("email",conent);
                    break;
                case 12:
                    break;
                case 13:
                    mapReal2.put("bankId",conent);
                    break;
                case 14:
                    break;
                case 15:
                    mapReal2.put("accountName",conent);
                    break;
                case 16:
                    mapReal2.put("bankAccountNumber",conent);
                    break;
                case 17:
                    mapReal.put("qrCode",conent);
                    mapReal2.put("qrCode",conent);
                    mapReal3.put("qrCode",conent);
                    break;
                case 18:
                    mapReal.put("location",conent);
                    break;
                case 19:
                    break;
                case 20:
                    break;
                case 21:
                    break;
                case 22:
                    break;
                case 23:
                    mapReal.put("salesPerson",conent);
                    break;
                case 24:
                    mapReal2.put("mdr",conent);
                    break;
                case 25:
                    break;
            }
        }
    }
    /**
     * 根据不同情况获取Java类型值
     * <ul><li>空白类型<ul><li>返回空字符串</li></ul></li></ul><ul><li>布尔类型</li><ul><li>返回Boulean类型值</li></ul></ul><ul><li>错误类型</li><ul><li>返回String类型值：Bad value</li></ul></ul><ul><li>数字类型</li><ul><li>日期类型</li><ul><li>返回格式化后的String类型，e.g.2017-03-15 22:22:22</li></ul><li>数字类型</li><ul><li>返回经过处理的java中的数字字符串，e.g.1.23E3==>1230</li></ul></ul> </ul><ul><li>公式类型</li><ul><li>公式正常</li><ul><li>返回计算后的String类型结果</li></ul></ul><ul><li>公式异常</li><ul><li>返回错误码，e.g.#DIV/0!；#NAME?；#VALUE!</li></ul></ul> </ul><ul><li>字符串类型</li><ul><li>返回String类型值</li></ul></ul>
     *
     * @param cell
     *            XSSFCell类型单元格
     * @return 返回Object类型值
     * @since 2017-03-26 00:05:36{@link #()}
     */
    public static Object getJavaValue(XSSFCell cell) {
        Object o = null;
        if(cell==null){
            return null;
        }
        int cellType = cell.getCellType();
        switch (cellType) {
            case XSSFCell.CELL_TYPE_BLANK:
                o = "";
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                o = cell.getBooleanCellValue();
                break;
            case XSSFCell.CELL_TYPE_ERROR:
                o = "Bad value!";
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                o = getValueOfNumericCell(cell);
                break;
            case XSSFCell.CELL_TYPE_FORMULA:
                try {
                    o = getValueOfNumericCell(cell);
                } catch (IllegalStateException e) {
                    try {
                        o = cell.getRichStringCellValue().toString();
                    } catch (IllegalStateException e2) {
                        o = cell.getErrorCellString();
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
                break;
            default:
                o = cell.getRichStringCellValue().getString();
        }
        return o;
    }
    // 获取数字类型的cell值
    private static Object getValueOfNumericCell(XSSFCell cell) {
        Boolean isDate = DateUtil.isCellDateFormatted(cell);
        Double d = cell.getNumericCellValue();
        Object o = null;
        if (isDate) {
            o = DateFormat.getDateTimeInstance()
                    .format(cell.getDateCellValue());
        } else {
            o = getRealStringValueOfDouble(d);
        }
        return o;
    }
    // 处理科学计数法与普通计数法的字符串显示，尽最大努力保持精度
    private static String getRealStringValueOfDouble(Double d) {
        String doubleStr = d.toString();
        boolean b = doubleStr.contains("E");
        int indexOfPoint = doubleStr.indexOf('.');
        if (b) {
            int indexOfE = doubleStr.indexOf('E');
            // 小数部分
            BigInteger xs = new BigInteger(doubleStr.substring(indexOfPoint
                    + BigInteger.ONE.intValue(), indexOfE));
            // 指数
            int pow = Integer.valueOf(doubleStr.substring(indexOfE
                    + BigInteger.ONE.intValue()));
            int xsLen = xs.toByteArray().length;
            int scale = xsLen - pow > 0 ? xsLen - pow : 0;
            doubleStr = String.format("%." + scale + "f", d);
        } else {
            java.util.regex.Pattern p = Pattern.compile(".0$");
            java.util.regex.Matcher m = p.matcher(doubleStr);
            if (m.find()) {
                doubleStr = doubleStr.replace(".0", "");
            }
        }
        return doubleStr;
    }

    public static List<Map<Object, Object>> getList1() {
        return list1;
    }

    public static void setList1(List<Map<Object, Object>> list1) {
        ExcelnputXlsx.list1 = list1;
    }

    public static List<Map<Object, Object>> getList2() {
        return list2;
    }

    public static void setList2(List<Map<Object, Object>> list2) {
        ExcelnputXlsx.list2 = list2;
    }

    public static List<Map<Object, Object>> getList3() {
        return list3;
    }

    public static void setList3(List<Map<Object, Object>> list3) {
        ExcelnputXlsx.list3 = list3;
    }
}
