package org.wxd.excel.utils;



import org.apache.commons.lang3.StringUtils;

import java.util.Collection;
import java.util.Iterator;
import java.util.Map;

/**
 * Created by wangxd on 2015/10/30.
 */
public class Assert {
    private Assert() {
    }

    public static void isTrue(boolean expression, String message) {
        if(!expression) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void isTrue(boolean expression) {
        isTrue(expression, "[Assertion failed] - this expression must be true");
    }

    public static void isFalse(boolean expression, String message) {
        if(expression) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void isFalse(boolean expression) {
        isFalse(expression, "[Assertion failed] - this expression must be false");
    }

    public static void isNull(Object object, String message) {
        if(object != null) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void isNull(Object object) {
        isNull(object, "[Assertion failed] - the object argument must be null");
    }

    public static void notNull(Object object, String message) {
        if(object == null) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void notNull(Object object) {
        notNull(object, "[Assertion failed] - this argument is required; it must not be null");
    }

    public static void isEmpty(CharSequence text, String message) {
        if(StringUtils.isNotEmpty(text)) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void isEmpty(CharSequence text) {
        isEmpty(text, "[Assertion failed] - this CharSequence argument must contains one character at least");
    }

    public static void notEmpty(CharSequence text, String message) {
        if(StringUtils.isEmpty(text)) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void notEmpty(CharSequence text) {
        notEmpty(text, "[Assertion failed] - this CharSequence argument must contains one character at least");
    }

    public static void isEmpty(Object[] array, String message) {
        if(array != null && array.length > 0) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void isEmpty(Object[] array) {
        isEmpty(array, "[Assertion failed] - this array must be empty: it must not contain any element");
    }

    public static void notEmpty(Object[] array, String message) {
        if(array == null || array.length == 0) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void notEmpty(Object[] array) {
        notEmpty(array, "[Assertion failed] - this array must not be empty: it must contain at least 1 element");
    }

    public static void isEmpty(Collection collection, String message) {
        if(collection != null && collection.size() > 0) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void isEmpty(Collection collection) {
        isEmpty(collection, "[Assertion failed] - this array must have no elements");
    }

    public static void notEmpty(Collection collection, String message) {
        if(collection == null || collection.isEmpty()) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void notEmpty(Collection collection) {
        notEmpty(collection, "[Assertion failed] - this collection must contain at least 1 element");
    }

    public static void isEmpty(Map map, String message) {
        if(map != null && map.size() > 0) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void isEmpty(Map map) {
        isEmpty(map, "[Assertion failed] - this array must be null or empty");
    }

    public static void notEmpty(Map map, String message) {
        if(map == null || map.isEmpty()) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void notEmpty(Map map) {
        notEmpty(map, "[Assertion failed] - this map must have at least one entry");
    }

    public static void isBlank(CharSequence text, String message) {
        if(StringUtils.isNotBlank(text)) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void isBlank(CharSequence text) {
        isBlank(text, "[Assertion failed] - this CharSequence argument must be null or blank");
    }

    public static void notBlank(CharSequence text, String message) {
        if(StringUtils.isBlank(text)) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void notBlank(CharSequence text) {
        notBlank(text, "[Assertion failed] - this CharSequence argument must not be null or blank");
    }

    public static void containsText(String textToSearch, String substring, String message) {
        if(!StringUtils.isEmpty(textToSearch) && !StringUtils.isEmpty(substring) && textToSearch.indexOf(substring) == -1) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void containsText(String textToSearch, String substring) {
        containsText(textToSearch, substring, "[Assertion failed] - this String argument must contain the substring [" + substring + "]");
    }

    public static void notContainsText(String textToSearch, String substring, String message) {
        if(!StringUtils.isEmpty(textToSearch) && !StringUtils.isEmpty(substring) && textToSearch.indexOf(substring) != -1) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void notContainsText(String textToSearch, String substring) {
        notContainsText(textToSearch, substring, "[Assertion failed] - this String argument must not contain the substring [" + substring + "]");
    }

    public static void startsWithText(String textToSearch, String substring, String message) {
        if(!StringUtils.isEmpty(textToSearch) && !StringUtils.isEmpty(substring) && !textToSearch.startsWith(substring)) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void startsWithText(String textToSearch, String substring) {
        startsWithText(textToSearch, substring, "[Assertion failed] - this String argument must start with the substring [" + substring + "]");
    }

    public static void notStartsWithText(String textToSearch, String substring, String message) {
        if(!StringUtils.isEmpty(textToSearch) && !StringUtils.isEmpty(substring) && textToSearch.startsWith(substring)) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void notStartsWithText(String textToSearch, String substring) {
        notStartsWithText(textToSearch, substring, "[Assertion failed] - this String argument must not start with the substring [" + substring + "]");
    }

    public static void endsWithText(String textToSearch, String substring, String message) {
        if(!StringUtils.isEmpty(textToSearch) && !StringUtils.isEmpty(substring) && !textToSearch.endsWith(substring)) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void endsWithText(String textToSearch, String substring) {
        endsWithText(textToSearch, substring, "[Assertion failed] - this String argument must end with the substring [" + substring + "]");
    }

    public static void notEndsWithText(String textToSearch, String substring, String message) {
        if(!StringUtils.isEmpty(textToSearch) && !StringUtils.isEmpty(substring) && textToSearch.endsWith(substring)) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void notEndsWithText(String textToSearch, String substring) {
        notEndsWithText(textToSearch, substring, "[Assertion failed] - this String argument must not end with the substring [" + substring + "]");
    }

    public static void noNullElements(Object[] array, String message) {
        if(array != null) {
            Object[] arr$ = array;
            int len$ = array.length;

            for(int i$ = 0; i$ < len$; ++i$) {
                Object each = arr$[i$];
                notNull(each, message);
            }

        }
    }

    public static void noNullElements(Object[] array) {
        noNullElements(array, "[Assertion failed] - this array must not contain any null elements");
    }

    public static void noNullElements(Collection collection, String message) {
        if(collection != null) {
            Iterator i$ = collection.iterator();

            while(i$.hasNext()) {
                Object each = i$.next();
                notNull(each, message);
            }

        }
    }

    public static void noNullElements(Collection collection) {
        noNullElements(collection, "[Assertion failed] - this array must not contain any null elements");
    }

    public static void noNullElements(Map map, String message) {
        if(map != null) {
            Iterator i$ = map.values().iterator();

            while(i$.hasNext()) {
                Object each = i$.next();
                notNull(each, message);
            }

        }
    }

    public static void noNullElements(Map map) {
        noNullElements(map, "[Assertion failed] - this map must not contain any null elements");
    }

    public static void isInstanceOf(Class type, Object obj, String message) {
        notNull(type, "Type to check against must not be null");
        if(!type.isInstance(obj)) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void isInstanceOf(Class type, Object obj) {
        isInstanceOf(type, obj, "Object of class [" + (obj != null?obj.getClass().getName():"null") + "] must be an instance of " + type);
    }

    public static void isAssignableFrom(Class superType, Class subType, String message) {
        notNull(superType, "Type to check against must not be null");
        if(subType == null || !superType.isAssignableFrom(subType)) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void isAssignableFrom(Class superType, Class subType) {
        isAssignableFrom(superType, subType, subType + " must be assignable to " + superType);
    }
}
