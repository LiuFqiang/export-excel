package annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
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
}