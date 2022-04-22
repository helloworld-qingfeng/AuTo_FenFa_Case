import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

@SuppressWarnings("unchecked")
public class utils {

    /*
       1.获取 目前的年 月
    */

    public static String Get_year_mouth(){
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String year_mouth = new SimpleDateFormat("yyyy-MM").format(new Date()).toString();
        return year_mouth;
    };


   /*
      2.获取当前时间，年月日时分秒
    */
   public static String Get_now_time(){
       return null;
   };



   /*
      3.获取当前java程序存放目录的目录地址;
    */
   public static String Get_Pro_now_dir(){
       File directory = new File("");//设定为当前文件夹
       return directory.getAbsolutePath();
   }



   /*
      4.根据Get_Pro_now_dir函数获取的目录地址，并且接受xlsx文件名称，拼接成xlsx全路径;
    */
   public static String Get_dir_Excel(String dirname,String excelname){
       File file = new File(dirname+"\\"+excelname);
       if(file.isFile()){
           //判断文件是否存在;
           return dirname+"\\"+excelname;
       }else {
         return null;
       }
   }


   /*
       5.获取当前java程序存放目录的父目录，即上一级目录
    */
   public static String  Get_Pro_now_FatherDir(String dirname){
       String Father_Dir_name = "";
       String[] split = dirname.split("\\\\");
       for(int i=0;i<split.length-1;i++){
           Father_Dir_name+=split[i]+"\\";
       }
       return Father_Dir_name;
   }



}
