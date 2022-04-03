package annotation;

import enums.DataType;
import enums.KeyType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ReportKey {

    String name() default "default";

    String description() default "default";

    KeyType keyType() default KeyType.SINGLE;

    DataType type() default DataType.TEXT;

    boolean temporary() default false;

    String dateFormatPattern() default "dd.MM.yyyy";

    String timeFormatPattern() default "HH:mm:ss";

}
