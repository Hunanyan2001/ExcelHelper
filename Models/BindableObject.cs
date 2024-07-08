using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelHelper.Models
{
    public class BindableObject : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        private readonly Dictionary<string, List<Action>> _propertyListeners = new Dictionary<string, List<Action>>();

        protected void OnPropertyChanged(string propertyName, Action callback)
        {
            if (!_propertyListeners.ContainsKey(propertyName))
                _propertyListeners.Add(propertyName, new List<Action>());

            _propertyListeners[propertyName].Add(callback);
        }

        protected void OnPropertyChanged(string propertyName)
        {
            if (_propertyListeners.ContainsKey(propertyName))
                foreach (var listener in _propertyListeners[propertyName])
                    listener();

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected virtual void Set<T>(ref T field, T newValue, [CallerMemberName] string? propertyName = null)
        {
            if (!EqualityComparer<T>.Default.Equals(field, newValue))
            {
                field = newValue;
                OnPropertyChanged(propertyName);
            }
        }
    }
}
