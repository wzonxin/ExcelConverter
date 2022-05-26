using System.Collections.Generic;

namespace ExcelConverter
{
    class MemPoolMgr : Singleton<MemPoolMgr>
    {
        private List<MemPoolBase> m_poolList = new List<MemPoolBase>();
        public void RegPool(MemPoolBase pool)
        {
            m_poolList.Add(pool);
        }

        public void ClearAllPool()
        {
            for (var i = 0; i < m_poolList.Count; i++)
            {
                m_poolList[i].ClearCache();
            }

            m_poolList.Clear();
        }
    }

    interface MemPoolBase
    {
        void ClearCache();
    }

    class MemPool<T> : Singleton<MemPool<T>> ,MemPoolBase where T : IMemPoolItem, new()
    {
        protected List<T> m_poolObjList = new List<T>();
        public MemPool()
        {
            MemPoolMgr.Instance.RegPool(this);
        }

        public static T New()
        {
            return Instance.Alloc();;
        }

        public static void Recycle(T obj)
        {
            Instance.Free(obj);
        }
        
        private T Alloc()
        {
            if (m_poolObjList.Count > 0)
            {
                var obj = m_poolObjList[0];
                m_poolObjList.RemoveAt(0);
                return obj;
            }

            T t = new T();
            t.Init();
            return t;
        }

        private void Free(T obj)
        {
            obj.OnDestroy();
            m_poolObjList.Add(obj);
        }

        public void ClearCache()
        {
            m_poolObjList.Clear();
        }
    }

    abstract class IMemPoolItem
    {
        public void Init(){}
        public abstract void OnDestroy();
    }
}
