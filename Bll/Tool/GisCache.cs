using System;
using System.Web;
using System.Collections;

namespace Bll.Tools
{
    public class GisCache
    {

        private static readonly object LockObj = new object();

        /// <summary>
        /// 获取数据缓存
        /// </summary>
        /// <param name="CacheKey">键</param>
        public static object GetCache(string CacheKey)
        {
            lock (LockObj)
            {
                System.Web.Caching.Cache objCache = HttpRuntime.Cache;
                return objCache[CacheKey];
            }
        }

        /// <summary>
        /// 设置数据缓存
        /// </summary>
        /// <param name="CacheKey">键</param>
        /// <param name="objObject">值</param>
        public static void SetCache(string CacheKey, object objObject)
        {
            System.Web.Caching.Cache objCache = HttpRuntime.Cache;
            objCache.Insert(CacheKey, objObject);
        }


        /// <summary>
        /// 设置数据缓存
        /// </summary>
        /// <param name="CacheKey">键</param>
        /// <param name="objObject">值</param>
        /// <param name="Timeout">有效时间</param>
        public static void SetCache(string CacheKey, object objObject, TimeSpan Timeout)
        {
            System.Web.Caching.Cache objCache = HttpRuntime.Cache;
            objCache.Insert(CacheKey, objObject, null, DateTime.MaxValue, Timeout, System.Web.Caching.CacheItemPriority.NotRemovable, null);
        }


        /// <summary>
        /// 设置数据缓存
        /// </summary>
        public static void SetCache(string CacheKey, object objObject, DateTime absoluteExpiration, TimeSpan slidingExpiration)
        {
            System.Web.Caching.Cache objCache = HttpRuntime.Cache;
            objCache.Insert(CacheKey, objObject, null, absoluteExpiration, slidingExpiration);
        }


        /// <summary>
        /// 更新数据缓存
        /// </summary>
        /// <param name="CacheKey">键</param>
        /// <param name="objObject">替换新值</param>
        public static void UpdateCahe(string CacheKey, object objObject)
        {
            var objCache = GetCache(CacheKey);
            if (null != objCache)
            {
                RemoveCache(CacheKey);
            }

            SetCache(CacheKey, objObject);
        }


        /// <summary>
        /// 移除指定数据缓存
        /// </summary>
        /// <param name="CacheKey">键</param>
        public static void RemoveCache(string CacheKey)
        {
            System.Web.Caching.Cache _cache = HttpRuntime.Cache;
            _cache.Remove(CacheKey);
        }

        /// <summary>
        /// 移除全部缓存
        /// </summary>
        public static void RemoveAllCache()
        {
            System.Web.Caching.Cache _cache = HttpRuntime.Cache;
            IDictionaryEnumerator CacheEnum = _cache.GetEnumerator();
            while (CacheEnum.MoveNext())
            {
                _cache.Remove(CacheEnum.Key.ToString());
            }
        }

    }
}
