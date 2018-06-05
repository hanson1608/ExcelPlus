/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: TypeReference
 * Author:   hanson
 * Date:   2018/5/27 15:46
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.excel.util;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.HashMap;
import java.util.Map;

/**
 * @author hanson
 * @since 1.0.0
 * 2018/5/27
 */
@SuppressWarnings("unused")
public class TypeReference<T> {
    private final Type type;
    private volatile Constructor<?> constructor;

    protected TypeReference() {
        Type superClass = getClass().getGenericSuperclass();
        this.type = ((ParameterizedType)superClass).getActualTypeArguments()[0];
    }

    /**
     * 取出所有的泛型到map里 泛型名：对应的真实的class
     * @return Map<String,Class<?>> 取出所有的泛型到map里 泛型名：对应的真实的class
     */
    public  Map<String,Class<?>> makeGenericMap(){
        Map<String,Class<?>> map = new HashMap<>(5);
        Type t = getClass().getGenericSuperclass();
        ParameterizedType p = (ParameterizedType) t;

        //rawType的参数是T里面包含的泛型的名称，如T,K,V等
        Type[] types = getRawType().getTypeParameters();
        Type actualType = p.getActualTypeArguments()[0];
        //判断T本身是否是包含泛型
        if(actualType instanceof ParameterizedType ){
            //这里获取的是T里泛型参数的实际参数类
            Type[] trueTypes = ((ParameterizedType)actualType).getActualTypeArguments();
            for (int i=0;i<types.length;i++){
                map.put(types[i].getTypeName(),(Class<?>)trueTypes[i]);
            }
        }
        return map;
    }

    /**
     * Instantiates a new instance of {@code T} using the default, no-arg
     * constructor.
     */
    @SuppressWarnings("unchecked")
    public T newInstance()
            throws NoSuchMethodException, IllegalAccessException,
            InvocationTargetException, InstantiationException {
        if (constructor == null) {
            Class<?> rawType = getRawType();
            constructor = rawType.getConstructor();
        }
        return (T) constructor.newInstance();
    }

    /**
     * Gets the referenced type.
     */
    public Type getType() {
        return this.type;
    }

    /**
     * 获取这个类的真实类
     * @return Class<?>
     */
    public Class<?> getRawType(){
        return type instanceof Class<?>
                ? (Class<?>) type
                : (Class<?>) ((ParameterizedType) type).getRawType();
    }
}
