import java.io.IOException;
import java.util.List;

public interface Form {

    abstract String getPlaceholder();
    abstract String getSeparator();

     /*
      1.读取xls文件，并且返回集合：
           filename：文件名称;
           sheetname：工作簿名字;
           array_fileds：字段数组;
      3.将指定的xlsx表中指定的sheet工作表取出,返回一个list<String>集合;
      4.在xlsx表中原有字段基础上,增加【sheet表名称】字段
     */

    abstract List<?> readxls(String filename, String sheetname, List<Short> array_fileds);


      /*
       4. 传入指定的list集合数据，对list集合中的每个元素进行过滤，返回过滤后的数据list集合;
          双列匹配过滤
          int是列号，即第几列；是用来作为条件查找的列，比如这个列的某个行数据，匹配到对应的条件即可
          condition是条件集合,比如第三字段的某个单元格是这个条件；人物名称
          String Separator  分隔符;，形成单元格的重要依据
      */
    abstract List<?> Filterdata(List<String> list, Short fieldnumber1, List<String> condition,String Heard_filed);


      /*
       6.读取一个xls表格，只读取他的字段头即可；返回一个list集合;
      */
    abstract List<?> read_field_name(String filename, String sheetname, String placeholder);


      /*
        7.读取一个xls表格，只读取他的某个字段[1个字段]；返回一个list集合;
        作为过滤条件
      */
    abstract List<?> read_filed(List<String> list, int fieldnumber1);


      /*
       8.wirte一个xlsx文件;
       8.1 List<List<String>> data 总数据
       8.2 List<String> name 人员姓名
      */
    abstract void write_xlsx(List<List<String>> data, List<String> name,String Dir_name_list,List<String> Dir_File_name_list) throws IOException;
}
