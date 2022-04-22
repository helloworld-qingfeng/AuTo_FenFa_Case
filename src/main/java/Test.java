import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

@SuppressWarnings("unchecked")
public class Test {
    public static void main(String[] args) throws IOException {

        //配置文件
        String ConfFile = "profile";
        //配置文件全路径
        String ConfDirFile = utils.Get_Pro_now_dir() + "\\" + ConfFile;


        //判断配置文件是否存在
        if (new File(ConfDirFile).isFile()) {

            //存在

           /*
               1.读取配置文件;
           */
            Properties prop = new Properties(); //创建流;
            File file = new File(ConfDirFile);
            InputStream inputStream = new FileInputStream(file);
            BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(inputStream,"UTF-8"));
            prop.load(bufferedReader);

                      //获取员工姓名列集合
                      String staff_name = (String) prop.get("staff_name");
                      String[] staff_name_split = staff_name.replaceAll("，", ",").split(",");
                      List<String> read_filed = new ArrayList<>();
                      for(String str:staff_name_split){
                          read_filed.add(str);
                      }


                      //获取xlsx文件名称(不是全路径)
                      String xlsx_name = (String) prop.get("Xlsx_name");

                     //员工需要打印的列，就是员工能看到的列
                     String xlsx_fileds = (String) prop.get("fileds");
                         //封装成集合;
                         List<Short> arry_fileds = new ArrayList<>();
                         String[] xlsx_fileds_split = xlsx_fileds.replaceAll("，", ",").split(",");
                         for(int i=0;i<=xlsx_fileds_split.length-1;i++){
                             Short y = Short.parseShort(xlsx_fileds_split[i]);
                             arry_fileds.add(y);
                         }

                     //获取条件列
                     String ob = (String) prop.get("IFfiled");
                     short IFfiled = Short.parseShort(ob);

                     //获取字段头;
                     String ob2 = (String) prop.get("Heard_filed");
                     String Heard_filed = ob2.replaceAll(",", "\t");

                     //获取员工部门的目录
                     String staff_dir = (String) prop.get("staff_dir");

            //关闭配置文件流
            bufferedReader.close();
            inputStream.close();

            //拼接xlsx文件的全路径
            String Xlsx_Dir_file = utils.Get_Pro_now_dir() + "\\" + xlsx_name;
            //获取sheet名称
            String sheet_name = utils.Get_year_mouth();
            //获得上一级目录
            String Get_Pro_now_FatherDir = utils.Get_Pro_now_FatherDir(utils.Get_Pro_now_dir());

            if (new File(Xlsx_Dir_file).isFile()) {
                //存在
                //多态创建
                Form form = new xlsx("\t","null");
                //读取表格，获取指定的列数据
                List<String> readxls = (List<String>) form.readxls(Xlsx_Dir_file, "Sheet1",arry_fileds);

                //依据员工姓名进行匹配，获得员工与案件列表的集合
                List<List<String>> filterdata = (List<List<String>>) form.Filterdata(readxls, IFfiled, read_filed,Heard_filed);

                //封装在职员工的目录文件点
                String ZaiZhi_YuanGong_Dir_list = Get_Pro_now_FatherDir+staff_dir;

                List<String> ZaiZhi_YuanGong_Dir_File_list = new ArrayList<>(); //文件点
                for(String str:read_filed){
                    ZaiZhi_YuanGong_Dir_File_list.add(Get_Pro_now_FatherDir+staff_dir+"\\"+str+".xlsx");

                }

                //开始写输入数据
                if(filterdata.size() == read_filed.size()){
                    form.write_xlsx(filterdata,read_filed,ZaiZhi_YuanGong_Dir_list,ZaiZhi_YuanGong_Dir_File_list);
                }
            }
        }
    }
}
