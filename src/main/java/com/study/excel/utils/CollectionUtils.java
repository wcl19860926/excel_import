package  com.study.excel.utils;



import java.util.*;
import java.util.function.Function;
import java.util.function.Supplier;

/**
 * 集合工具类。
 */
public  class CollectionUtils {
	/**
	 * 判断指定的集合是否为空。
	 * 
	 * @param collection
	 *            待判断的集合
	 * @return 返回指定的集合是否为空。
	 */
	public static Boolean isEmpty(Collection<?> collection) {
		return (collection == null || collection.isEmpty());
	}

	/**
	 * 判断指定的集合是否不为空。
	 * 
	 * @param collection
	 *            待判断的集合
	 * @return 返回指定的集合是否不为空。
	 */
	public static Boolean isNotEmpty(Collection<?> collection) {
		return !isEmpty(collection);
	}

	/**
	 * 判断指定的数组是否为空。
	 * 
	 * @param array
	 *            待判断的数组
	 * @return 返回指定的数组是否为空。
	 */
	public static Boolean isEmpty(Object[] array) {
		return (array == null || array.length == 0);
	}

	/**
	 * 判断指定的数组是否不为空。
	 * 
	 * @param array
	 *            待判断的数组
	 * @return 返回指定的数组是否不为空。
	 */
	public static Boolean isNotEmpty(Object[] array) {
		return !isEmpty(array);
	}

	/**
	 * 判断指定的Map是否为空。
	 * 
	 * @param map
	 *            待判断的Map
	 * @return 返回指定的Map是否为空。
	 */
	public static Boolean isEmpty(Map<?, ?> map) {
		return (map == null || map.isEmpty());
	}

	/**
	 * 判断指定的Map是否不为空。
	 * 
	 * @param map
	 *            待判断的Map
	 * @return 返回指定的Map是否不为空。
	 */
	public static Boolean isNotEmpty(Map<?, ?> map) {
		return !isEmpty(map);
	}

	/**
	 * 移除List中重复的元素。
	 * 
	 * @param <T>
	 *            元素类型
	 * 
	 * @param list
	 *            列表
	 */
	public static <T> void removeDuplicate(List<T> list) {
		HashSet<T> set = new HashSet<T>(list);
		list.clear();
		list.addAll(set);
	}

	/**
	 * 判断数组中是否包含指定元素。
	 * 
	 * @param <T>
	 *            数组类型
	 * @param elements
	 *            数组
	 * @param elementToFind
	 *            待查找的元素
	 * @return 如果数组中包含指定元素返回true，否则返回false。
	 */
	public static <T> Boolean contains(T[] elements, T elementToFind) {
		List<T> elementsList = Arrays.asList(elements);
		return elementsList.contains(elementToFind);
	}


	/**
	 *
	 * @param lists
	 * @param apply
	 * @param <U>
	 * @param <R>
	 * @return
	 */
	public static <U ,R>  List<List<R>>  getSplicedPropertiesList(List<U>  lists , Function<U ,R> apply , int  subSize){
           if(isEmpty(lists)){
           	return   new ArrayList<>();
		   }
		int sumSize =  lists.size();
           int tempSize =0;
           List<List<R>>    resultList  = new ArrayList<>();
           List<R>   subList  =  new ArrayList<>();
           for(U     u  :  lists){
			   tempSize++ ;
			   subList.add(apply.apply(u));
			   if(tempSize== subSize){
				   subList = new ArrayList<>();
				   resultList.add(subList);
				   tempSize=0;
			   }
		   }
          return resultList;
	}


	/**
	 *
	 * @param lists
	 * @param subSize
	 * @param <U>
	 * @return
	 */
	public static <U>  List<List<U>>  getSplicedPropertiesList(List<U>  lists , int  subSize){
		if(isEmpty(lists)){
			return   new ArrayList<>();
		}
		int sumSize =  lists.size();
		int tempSize =0;
		List<List<U>>    resultList  = new ArrayList<>();
		List<U>   subList  =  new ArrayList<>();
		for(U     u  :  lists){
			tempSize++ ;
			subList.add(u);
			if(tempSize== subSize){
				subList = new ArrayList<>();
				resultList.add(subList);
				tempSize=0;
			}
		}
		return resultList;
	}

	/**
	 *
	 * @param lists
	 * @param apply
	 * @param <U>
	 * @param <R>
	 * @return
	 */
	public static <U ,R>  List<R>  getPropertiesList(List<U>  lists , Function<U ,R> apply ){
		if(isEmpty(lists)){
			return   new ArrayList<>();
		}
		List<R>   propertiesList  =  new ArrayList<>();
		for(U     u  :  lists){
			propertiesList.add(apply.apply(u));
		}
		return propertiesList;
	}

	/**
	 *
	 * @param lists
	 * @param function
	 * @param <U>
	 * @return
	 */
	public static <U>  List<String>  getNotNullStringPropertiesList(List<U>  lists , ToStringFunction<U> function ){
		if(isEmpty(lists)){
			return   new ArrayList<>();
		}
		List<String>   propertiesList  =  new ArrayList<>();
		String  value = null;
		for(U     u  :  lists){
			value = function.applyAsString(u);
			if(StringUtil.isNotBlank(value)){
				propertiesList.add(value);
			}
		}
		return propertiesList;
	}

	/**
	 *
	 * @param lists
	 * @param function
	 * @param <K>
	 * @param <U>
	 * @return
	 */
	public static <K ,U>  Map<K ,U>  getPropertiesMap(List<U>  lists , Function<U ,K> function ){
		if(isEmpty(lists)){
			return   new HashMap<>();
		}
		Map<K, U>   propertiesMap  =  new HashMap<>();
		for(U    u  :  lists){
			propertiesMap.put(function.apply(u) , u);
		}
		return propertiesMap;
	}

	/**
	 *
	 * @param lists
	 * @param keyFunction
	 * @param valueFunction
	 * @param <U>
	 * @param <K>
	 * @param <V>
	 * @return
	 */
	public static <U ,K ,V>  Map<K ,V>  getPropertiesMap(List<U>  lists , Function<U ,K> keyFunction  ,Function<U ,V> valueFunction ){
		if(isEmpty(lists)){
			return   new HashMap<>();
		}
		Map<K, V>   propertiesMap  =  new HashMap<>();
		for(U    u  :  lists){
			propertiesMap.put(keyFunction.apply(u) , valueFunction.apply(u));
		}
		return propertiesMap;
	}
}
