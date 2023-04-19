package org.example;

import java.lang.reflect.Field;
import java.util.List;

public class CollectionUtils {
    public static void printCollection(List<?> collection) {
        // 获取对象的Class对象
        Class<?> clazz = collection.get(0).getClass();

        // 获取所有属性的Field对象
        Field[] fields = clazz.getDeclaredFields();

        // 输出属性名称
        StringBuilder sb = new StringBuilder();
        for (Field field : fields) {
            if (field.isAnnotationPresent(Property.class)) {
                Property property = field.getAnnotation(Property.class);
                sb.append(property.name()).append("\t");
            }
        }
        System.out.println(sb.toString());

        // 输出集合数据
        for (Object obj : collection) {
            sb = new StringBuilder();
            for (Field field : fields) {
                if (field.isAnnotationPresent(Property.class)) {
                    try {
                        Object value = field.get(obj);
                        sb.append(value).append("\t");
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
            }
            System.out.println(sb.toString());
        }
    }
}

