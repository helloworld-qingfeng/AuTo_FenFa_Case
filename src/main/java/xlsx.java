
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class xlsx extends utils implements Form {

    String placeholder = null; //占位符，null

    //Separator:分隔符;
    String Separator =null; //分隔符;

    public xlsx(String Separator,String placeholder) {
       this.placeholder=placeholder;
       this.Separator=Separator;
    }

    public xlsx(){};

    @Override
    public String getPlaceholder() {
        return this.placeholder;
    }

    @Override
    public String getSeparator() {
        return this.Separator;
    }

    @Override
    public List<?> readxls(String filename, String sheetname,List<Short> array_fileds) {
         /*
        1.读取xlsx文件中指定的sheet表，并且返回list集合，string
           */
        String path= filename;
        FileInputStream fileInputStream=null;
        XSSFWorkbook workbook=null;
        List<String> list = new ArrayList<>();
        try {
            fileInputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheet(sheetname);
            int maxRow = sheet.getLastRowNum();
            for(int i=1;i<maxRow;i++){
                XSSFRow row = sheet.getRow(i);
                String data = "";
                if(row!=null){
                    int maxRol = sheet.getRow(i).getLastCellNum();
                    for(int j=0;j<=array_fileds.size()-1;j++){
                        XSSFCell cell = row.getCell(array_fileds.get(j));
                        data+=cell+this.Separator;
                    }
                    list.add(data);
                }
            }
            return list;

        }catch (FileNotFoundException e){
            e.printStackTrace();
        }catch (IOException e){
            e.printStackTrace();
        }finally {
            try{
                workbook.close();
            }catch (IOException e){
                e.printStackTrace();
            }
        }
        return list;
    }

    @Override
    public List<?> Filterdata(List<String> list, Short fieldnumber1, List<String> condition,String Heard_filed) {
        //返回一个集合;
        List<List<String>> lists = new ArrayList<>();

        //判断2个集合是否为空；
        //判断列号是否大于0；
        //判断分隔符是否为空;
        if(list != null && condition != null && fieldnumber1 > 0 && this.Separator != null && this.Separator.length() > 0){

            //拆分条件列表
            for(String ChuLiRen: condition){

                //创建 某个条件列表 的集合
                List<String> ChuLiRen_list = new ArrayList<>();

                //添加头
                ChuLiRen_list.add(Heard_filed);

                //拆分总表数据;拆成行;
                for(int line_data=list.size()-1;line_data >=0;line_data--){

                    //提取行的某个列，和对应的条件进行对比
                    String[] split = list.get(line_data).split(this.Separator);


                    if (split.length > fieldnumber1-1) {

                        //判断对应的字段数据 不为空
                        if (split[fieldnumber1] != null && split[fieldnumber1].length() > 0 && !split[fieldnumber1].equals("null")) {
                            if(split[fieldnumber1].equals(ChuLiRen)){
                                //添加符合条件的数据行
                                ChuLiRen_list.add(list.get(line_data));
                                //删除总表里原有的数据行;
                                list.remove(line_data);
                            }
                        }

                    }
                }
                lists.add(ChuLiRen_list);
            }

        }
        return lists;
    }

    @Override
    public List<?> read_field_name(String filename, String sheetname, String placeholder) {
        return null;
    }

    @Override
    public List<?> read_filed(List<String> list, int fieldnumber1) {

        List<String> filed_list = new ArrayList<String>();

        //判断 字段号是否为0；判断分隔符是否为空；
        if (fieldnumber1 != 0 && this.Separator != null && this.Separator.length() > 0) {
            //分割
            for (String str : list) {
                String[] split = str.split(this.Separator);
                //判断行数据，是否有那么多字段
                if (split.length > fieldnumber1-1) {
                    //判断对应的字段数据 不为空
                    if (split[fieldnumber1] != null && split[fieldnumber1].length() > 0 && !split[fieldnumber1].equals("null")) {
                        //填装集合
                        filed_list.add(split[fieldnumber1]);
                    }
                }
            }
            //去重
            return filed_list.stream().distinct().collect(Collectors.toList());
        }
        return filed_list;
    }



    @Override
    public void write_xlsx(List<List<String>> data, List<String> name,String Dir_name_list,List<String> Dir_File_name_list) throws IOException {

        //判断参数是否合法
        //判断目录是否存在
        if(name.size()==data.size() && data != null && name != null && new File(Dir_name_list).isDirectory()){

            for(int i=0;i<=name.size()-1;i++){

                    //创建
                    Workbook wb = new XSSFWorkbook();

                    //创建对应的员工xlsx表;
                    OutputStream fileOut = new FileOutputStream(Dir_File_name_list.get(i));
                    //创建对应的员工sheet表;
                    Sheet sheet = wb.createSheet(name.get(i)+ utils.Get_year_mouth());

                    //设置专利名称那个字段的长度;
                    sheet.setColumnWidth((short)2,(short)12000);

                    //设置时间那个字段长度
                     sheet.setColumnWidth((short)1,(short)6000);

                    List<String> YuanGong_Data = data.get(i);   //提取 某个员工 数据集合

                    //写数据
                    for (int y=0;y<=YuanGong_Data.size()-1;y++) {  //提取行号

                        if(y==0){
                            //写字段头
                            Row row = sheet.createRow(y);  //行号;
                            String[] split = YuanGong_Data.get(y).split(this.Separator);  //将行数据，分裂成若干字段数据;

                            XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
                            //设置背景色
                            style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            //水平居中
                            style.setAlignment(HorizontalAlignment.CENTER);
                            for(int t=0;t<=split.length-1;t++){
                                Cell cell = row.createCell(t);
                                cell.setCellValue(split[t]);  //分裂成字段号
                                cell.setCellStyle(style);  //设置背景色
                            }

                        }else {
                            XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
                            style.setAlignment(HorizontalAlignment.CENTER);

                            //写数据本体
                            Row row = sheet.createRow(y);  //行号;
                            String[] split = YuanGong_Data.get(y).split(this.Separator);  //将行数据，分裂成若干字段数据;

                            for(int t=0;t<=split.length-1;t++){
                                Cell cell = row.createCell(t);

                                //检测是否是状态字段
                                if(t==3){
                                   if(!split[t].equals("已完成")){
                                       //不是已完成的
                                      XSSFFont font = (XSSFFont) wb.createFont();
                                      font.setColor(XSSFFont.COLOR_RED); //红色
                                      XSSFRichTextString ts= new XSSFRichTextString(split[t]);
                                      ts.applyFont(0,ts.length(),font); //0起始索引,2结束索引    标题长度
                                      cell.setCellValue(ts);      //分裂成字段号
                                      cell.setCellStyle(style);  //设置格式居中
                                   }else {
                                       cell.setCellValue(split[t]);  //分裂成字段号
                                       cell.setCellStyle(style);  //设置格式居中
                                   }
                                }else {
                                    //不是状态字段
                                    cell.setCellValue(split[t]);  //分裂成字段号
                                    cell.setCellStyle(style);  //设置格式居中
                                }
                            }
                        }
                    }
                    wb.write(fileOut);
                    fileOut.flush();
                    fileOut.close();
                    wb.close();
            }
        }
    }
}
