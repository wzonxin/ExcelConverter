using System;
using System.Collections.Generic;

namespace ExcelConverter
{
    class EventDispatcher
    {
        private static Dictionary<TaskType, Delegate> _dictEvent;
        private static List<ITask> _taskList;

        static EventDispatcher()
        {
            _dictEvent = new Dictionary<TaskType, Delegate>();
            _taskList = new List<ITask>();
        }

        public static void SendEvent(TaskType type)
        {
            Delegate action;
            if (_dictEvent.TryGetValue(type, out action))
            { 
                var task = new TimerTask();
                task.Action = action as Action;
                _taskList.Add(task);
            }
        }
        
        public static void SendEvent<T>(TaskType type, T t)
        {
            Delegate action;
            if (_dictEvent.TryGetValue(type, out action))
            { 
                var task = new TimerTask<T>();
                task.Data = t;
                task.Action = action as Action<T>;
                _taskList.Add(task);
            }
        }
        
        public static void SendEvent<T, K>(TaskType type, T t, K k)
        {
            Delegate action;
            if (_dictEvent.TryGetValue(type, out action))
            {
                var task = new TimerTask<T, K>();
                task.Data = t;
                task.Data2 = k;
                task.Action = action as Action<T, K>;
                _taskList.Add(task);
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
            if(_taskList.Count > 0)
            {
                var task = _taskList[0];
                task.DoAction();
                _taskList.RemoveAt(0);
            }
        }

        public static void Clear()
        {
            _dictEvent.Clear();
            _taskList.Clear();
        }
    }
}
