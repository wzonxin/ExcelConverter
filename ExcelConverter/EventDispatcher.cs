using System;
using System.Collections.Generic;

namespace ExcelConverter
{
    class EventDispatcher
    {
        private static Dictionary<TaskType, Delegate> _dictEvent;
        private static Queue<TaskBase> _taskList;
        private const int _roundHandleMaxEventCnt = 10;

        static EventDispatcher()
        {
            _dictEvent = new Dictionary<TaskType, Delegate>();
            _taskList = new Queue<TaskBase>();
        }

        public static void SendEvent(TaskType type)
        {
            Delegate action;
            if (_dictEvent.TryGetValue(type, out action))
            { 
                var task = MemPool<TimerTask>.New();
                task.Action = action as Action;
                _taskList.Enqueue(task);
            }
        }
        
        public static void SendEvent<T>(TaskType type, T t)
        {
            Delegate action;
            if (_dictEvent.TryGetValue(type, out action))
            { 
                var task = MemPool<TimerTask<T>>.New();
                task.Data = t;
                task.Action = action as Action<T>;
                _taskList.Enqueue(task);
            }
        }
        
        public static void SendEvent<T, K>(TaskType type, T t, K k)
        {
            Delegate action;
            if (_dictEvent.TryGetValue(type, out action))
            {
                var task = MemPool<TimerTask<T, K>>.New();
                task.Data = t;
                task.Data2 = k;
                task.Action = action as Action<T, K>;
                _taskList.Enqueue(task);
            }
        }

        public static void RegdEvent(TaskType type, Action vt)
        {
            _dictEvent.Add(type, vt);
        }
        
        public static void RegdEvent<T>(TaskType type, Action<T> vt)
        {
            _dictEvent.Add(type, vt);
        }
        
        public static void RegdEvent<T, K>(TaskType type, Action<T, K> vt)
        {
            _dictEvent.Add(type, vt);
        }

        public static void CheckTick()
        {
            int executeCnt = 0;
            while(_taskList.Count > 0 && executeCnt <= _roundHandleMaxEventCnt)
            {
                var task = _taskList.Dequeue();
                task.DoAction();

                task.Recycle();
                executeCnt++;
            }
        }

        public static void RemoveEvent(TaskType type)
        {
            _dictEvent.Remove(type);
        }

        public static void Clear()
        {
            _dictEvent.Clear();
            _taskList.Clear();
        }
    }
}
