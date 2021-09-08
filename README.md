# export-excel
Export excel using POI custom annotation @ Excel

自定义注解实现导出Excel

1、新增注解@Excel


public @interface Excel {

    /**
     * 表头名字
     */
    String name() default "";

    /**
     * 列顺序
     */
    int sort() default Integer.MAX_VALUE;

    /**
     * 日期格式化
     */
    String pattern() default "yyyy-MM-dd";

    /**
     * 列宽度
     */
    double width() default 16;

    /**
     * 默认值
     */
    String defaultValue() default "";

    /**
     * 对齐方式
     */
    Align align() default Align.CENTER;

    /**
     * 码值转换
     * 替换类型 {"0_女", "1_男"}
     */
    String[] codeValue() default {};

    enum Align {
        AUTO(0), LEFT(1), CENTER(2), RIGHT(3);
        private final int value;

        Align(int value) {
            this.value = value;
        }

        public int value() {
            return this.value;
        }
    }
    
2、添加导出工具类ExportExcelUtils

3、替换Application.java生成文件的路径为本地

4、调用
  
  InputStream is = ExportExcelUtil.export(studentDtoList, "sheet1", StudentDto.class);
  
  其中studentDtoList为List<StudentDto>
  
  生成Excel如下：
  ![image](https://img2020.cnblogs.com/blog/1597479/202109/1597479-20210908164349310-665405923.png)
  
